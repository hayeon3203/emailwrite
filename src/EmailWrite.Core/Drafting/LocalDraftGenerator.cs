using System.Diagnostics;
using System.Globalization;
using System.Runtime.Intrinsics.X86;
using System.Security.Cryptography;
using LLama;
using LLama.Common;
using LLama.Sampling;
using EmailWrite.Core.OutlookIntegration;

namespace EmailWrite.Core.Drafting;

/// <summary>
/// Runs draft generation locally via LLamaSharp against a HyperCLOVA X SEED 1.5B GGUF, downloaded
/// on demand into a machine-wide model cache. Mode-switch UI and custom model URLs are
/// still M9; the minimum spec (AVX2) is checked here.
/// </summary>
public sealed class LocalDraftGenerator : IDraftGenerator, IDisposable
{
    // HyperCLOVA X SEED 1.5B, not Mi:dm 2.3B. Mi:dm turned a request into a promise — "국문 버전도
    // 만들어달라고" came back as "만들어드리겠습니다" — 14 times in 20, and five separate fixes
    // failed to move it. The sentence reads fine, so it gets copied and sent meaning the opposite.
    // CLOVA is 8/20 on the same prompt and 0/20 with the by-direction examples that are on today.
    // Smaller (1.06GB vs 1.33GB) and faster (6.6s vs 10.3s on turn 1) as well.
    // The name and the URL have to move together: leave the URL on Mi:dm and a new PC downloads
    // Mi:dm's weights, saves them under CLOVA's name, and loads the wrong model with no error.
    private const string ModelFileName = "hyperclovax-seed-text-instruct-1.5b-q4_k_m.gguf";
    private const string ModelDownloadUrl =
        "https://huggingface.co/kexplo/HyperCLOVAX-SEED-Text-Instruct-1.5B-Q4_K_M-GGUF/resolve/main/hyperclovax-seed-text-instruct-1.5b-q4_k_m.gguf";
    // Matches Hugging Face's own LFS oid for that blob, which is what makes this worth having:
    // the file is one person's conversion in a repo they can force-push at any time, and nothing
    // upstream would tell us. Confirmed against a fresh 1,133,974,368-byte download (2026-08-14).
    // Bump this and ModelFileName together whenever the model moves.
    private const string ModelSha256 = "95659a99a087fe5803150f47569dd55716a6806b6c3bf18e901ca0d341a701e4";

    private readonly string _modelPath;
    private readonly SemaphoreSlim _initLock = new(1, 1);
    // llama.cpp's context is not reentrant: a second llama_decode entering while the first is
    // still running corrupts native memory and kills the process with an AccessViolationException
    // (not catchable — the whole app dies). Reachable from the UI just by hitting F9 twice while
    // a draft is still generating, so inference is serialized here rather than in each caller.
    private readonly SemaphoreSlim _inferLock = new(1, 1);
    private LLamaWeights? _weights;
    private ModelParams? _modelParams;

    // Prefill is essentially the entire cost of a turn: capping generation at 8 tokens instead of
    // 512 moved a turn from 14.0s to 13.6s, so ~13.6s of it is processing the prompt. Almost all of
    // that prompt is identical to the previous turn's — same system prompt, same examples, same
    // mail, same earlier messages — so an InteractiveExecutor holds the KV cache between turns and
    // only the newly appended text gets fed. That drops a revision turn to ~2.6s.
    //
    // The bookkeeping below tracks what the cache actually holds rather than re-rendering the
    // whole prompt and diffing: the rendered form of a past turn is not byte-identical to what was
    // generated (the reply gets trimmed and its control tokens stripped before it is shown), so a
    // prefix comparison breaks on a stray newline and silently falls back to a full prefill.
    private LLamaContext? _context;
    private InteractiveExecutor? _executor;
    // Read off the loaded model's own chat template once, never assumed. Null means the probe
    // could not make sense of it, and every turn then pays a full prefill rather than guessing.
    private TurnMarkers? _markers;
    // Which native CPU kernel llama.cpp actually loaded. Null until the weights are in.
    private string? _kernel;
    // The classifier runs on its own StatelessExecutor rather than the session context above:
    // feeding it through _executor would overwrite the KV cache the draft turns are built on, and
    // a full prefill per turn costs far more than this call saves. Stateless makes its own context
    // per inference from _kindParams (512 tokens, not 4096 — the prompt is one sentence), so there
    // is nothing to keep in sync and nothing to dispose between runs.
    private StatelessExecutor? _kindExecutor;
    private ModelParams? _kindParams;
    // How long the classification took, for the status line. Null when it did not run.
    private string? _lastClassify;
    // Exactly what the classifier said, before ParseKind. Kept because "판정불가" on its own does
    // not say whether the model refused, rambled, or answered a fourth word.
    private string? _lastRawKind;

    /// <summary>
    /// What the probe found, for DraftProbe to print when a new .gguf is being vetted — a model
    /// whose markers come back null still works, but silently at one full prefill per turn.
    /// Null until the weights have been loaded.
    /// </summary>
    internal TurnMarkers? Markers => _markers;
    // Which conversation the cache belongs to. -1 means "nothing usable in there". This used to
    // compare the EmailContext by reference, which worked only as long as there was always a mail:
    // with free-writing both sides are null, null equals null, and a brand new conversation would
    // look like a continuation of the last one.
    private int _sessionId = -1;
    private int _fedMessageCount;
    private bool _kvEndsAtTurnEnd;

    /// <summary>
    /// Where the "내려받는 중 42%" line goes while the model is being fetched. The first run on a
    /// fresh PC downloads 1.06GB before it can answer anything, and with nothing reporting it the
    /// popup sits on its typing indicator for minutes — indistinguishable from a hang. An empty
    /// report means "done, put the default text back".
    /// </summary>
    public IProgress<string>? Progress { get; set; }

    /// <summary>
    /// What the last run actually spent its time on: prefill (time to the first token) against
    /// decode (how many tokens came out). A single wall-clock number can't tell a long mail from a
    /// draft that ran to <c>MaxTokens</c>, and those want opposite fixes — the app was reported
    /// slower than every probe measurement with no way to see which one it was.
    /// </summary>
    public string? LastStats { get; private set; }

    /// <summary>
    /// What the classifier called the last opening turn, or null when it was off, skipped (a
    /// revision turn has no direction) or could not answer with one of the three words. Exposed so
    /// a 20-run measurement can separate "the label was wrong" from "the label did not help".
    /// </summary>
    public string? LastKind { get; private set; }

    public LocalDraftGenerator()
    {
        // Deliberately NOT %LocalAppData%: llama.cpp's gguf_init_from_file opens the model
        // with plain ANSI fopen, so a non-ASCII user profile name (e.g. C:\Users\조하연\...)
        // arrives mangled and the load fails. CommonApplicationData is C:\ProgramData —
        // always ASCII, writable by Users, and shared so a second user on the same PC
        // doesn't re-download 1.06GB.
        var modelDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "EmailWrite", "Models");
        Directory.CreateDirectory(modelDir);

        // EMAILWRITE_MODEL swaps the .gguf without a rebuild, so alternatives already sitting
        // in the cache can be A/B'd offline. Auto-download only knows the default model's URL;
        // an overridden file has to be there already.
        var fileName = Environment.GetEnvironmentVariable("EMAILWRITE_MODEL");
        _modelPath = Path.Combine(modelDir, string.IsNullOrWhiteSpace(fileName) ? ModelFileName : fileName);
    }

    public async Task<string> GenerateDraftAsync(
        EmailContext? emailContext,
        IReadOnlyList<ChatMessage> conversation,
        int sessionId = 0,
        CancellationToken cancellationToken = default)
    {
        // The bundled llama.cpp backends include a noavx build, but every measurement behind the
        // model choice was taken on AVX2 and the fallback is slow enough to be unusable for a
        // draft you wait on. Checked before anything else so a colleague on an old PC gets told
        // what is wrong, rather than a 1.06GB download followed by a native load failure.
        if (!Avx2.IsSupported)
        {
            // Was "ANTHROPIC_API_KEY 환경변수를 설정하면", which stopped being true once the key
            // moved into the popup: it sent a colleague off to set a machine environment variable
            // for a backend that is no longer the only API option, to fix something two clicks
            // away on the screen they are already looking at.
            return "이 PC의 CPU는 AVX2를 지원하지 않아 로컬 모델을 쓸 수 없어요. "
                 + "위 목록에서 Solar나 Claude를 고르면 API 모드로 쓸 수 있어요 (키가 없으면 입력칸이 같이 떠요).";
        }

        try
        {
            await EnsureWeightsAsync(cancellationToken);
            // Whatever the download/load reported is stale now: generation is next.
            Progress?.Report(string.Empty);
        }
        catch (Exception ex)
        {
            Progress?.Report(string.Empty);
            return $"로컬 모델을 준비하는 중 오류가 발생했어요(다운로드에 실패했을 수 있어요): {ex.Message}";
        }

        await _inferLock.WaitAsync(cancellationToken);
        try
        {
            // Continue the live cache only when this is the same conversation and it has merely
            // grown; a reopened popup arrives with a new session id and starts over. The mail is
            // fixed when the session opens, so the id alone settles both questions.
            string toFeed;
            if (_executor != null && _markers != null && _sessionId == sessionId && conversation.Count > _fedMessageCount)
            {
                toFeed = BuildContinuation(conversation.Skip(_fedMessageCount));
            }
            else
            {
                StartSession();
                _sessionId = sessionId;
                // Only the opening turn has a direction to establish: everything after it is a
                // revision request ("조금 더 부드럽게"), and the continuation path above never
                // rebuilds the prompt anyway. So this is at most one extra call per conversation —
                // plus one per "다시", which starts the opening turn over.
                LastKind = Environment.GetEnvironmentVariable("EMAILWRITE_KIND") == "1"
                        && conversation.Count == 1 && conversation[0].Role == ChatRole.User
                    ? await ClassifyAsync(conversation[0].Text, cancellationToken)
                    : null;
                toFeed = BuildPrompt(emailContext, conversation, LastKind);
            }

            var inferenceParams = new InferenceParams
            {
                MaxTokens = 512,
                // The model's own turn-end string, not a guessed one. This used to be ChatML's
                // "<|im_end|>", which Mi:dm never emits — the anti-prompt simply never fired.
                AntiPrompts = _markers == null ? new List<string>() : new List<string> { _markers.TurnEnd },
            };

            if (ReadTemperature() is { } temperature)
            {
                inferenceParams.SamplingPipeline = new DefaultSamplingPipeline { Temperature = temperature };
            }

            var text = new System.Text.StringBuilder();
            var watch = Stopwatch.StartNew();
            var prefill = TimeSpan.Zero;
            var tokens = 0;
            await foreach (var token in _executor!.InferAsync(toFeed, inferenceParams, cancellationToken))
            {
                if (tokens++ == 0)
                {
                    prefill = watch.Elapsed;
                }

                text.Append(token);
            }

            // 512 is MaxTokens: seeing it here means the model never emitted its turn-end and the
            // draft was cut off, which costs the whole cap in decode time.
            LastStats = $"커널 {_kernel} · 프리필 {prefill.TotalSeconds:F1}s · {tokens}토큰 "
                      + $"{tokens / Math.Max(0.1, (watch.Elapsed - prefill).TotalSeconds):F1}t/s "
                      + $"· 프롬프트 {toFeed.Length}자"
                      + (_lastClassify == null ? "" : $" · 성격 {LastKind ?? "판정불가"} {_lastClassify}");

            // The cache now also holds the reply just generated, which the caller will hand back
            // as the next conversation entry — hence +1.
            var raw = text.ToString();
            _fedMessageCount = conversation.Count + 1;
            _kvEndsAtTurnEnd = _markers != null
                && raw.TrimEnd().EndsWith(_markers.TurnEnd.TrimEnd(), StringComparison.Ordinal);

            var result = StripThinking(StripTurnMarkers(raw, _markers));
            return string.IsNullOrWhiteSpace(result) ? "응답을 받지 못했어요." : result;
        }
        catch (Exception ex)
        {
            // A failed or cancelled run leaves the KV cache holding an unknown amount of the
            // prompt. Forget the fast path rather than feed the next turn a continuation that
            // assumes a prefix which may not be there.
            _sessionId = -1;
            return $"로컬 모델 추론 중 오류가 발생했어요: {ex.Message}";
        }
        finally
        {
            _inferLock.Release();
        }
    }

    /// <summary>
    /// Asks the model one question about one sentence: is this a request, a promise, or neither.
    /// The answer becomes the 성격 field of the draft prompt (<see cref="DraftPromptBuilder"/>).
    ///
    /// The reason this can work while the same model gets the direction wrong when drafting is that
    /// this call sees *only* the user's sentence — no mail, no thread history, no signature block.
    /// §9 traced the inversion to exactly that material ("상대방 문장 → 내 문장"), and none of it
    /// is in this prompt. Doing the same judgement inside the draft call would throw that away.
    ///
    /// The draft path calls this only when EMAILWRITE_KIND=1, so the measured baseline stays what
    /// runs by default and the label can be A/B'd without a rebuild — same pattern as
    /// EMAILWRITE_WHOAMI/EX2/NO_EX_ASK. Any failure returns null, which builds today's prompt
    /// exactly: a wrong label points the draft the wrong way, so falling back to "no label" is the
    /// only safe direction to fail in.
    /// </summary>
    private async Task<string?> ClassifyAsync(string intent, CancellationToken cancellationToken)
    {
        _lastClassify = null;
        try
        {
            _kindExecutor ??= new StatelessExecutor(_weights!, _kindParams!);

            var template = new LLamaTemplate(_weights!) { AddAssistant = true };
            template.Add("system", DraftPromptBuilder.KindSystemPrompt);
            foreach (var (role, text) in DraftPromptBuilder.KindExampleTurns)
            {
                template.Add(role == ChatRole.User ? "user" : "assistant", text);
            }

            template.Add("user", DraftPromptBuilder.KindQuestion(intent));
            var prompt = System.Text.Encoding.UTF8.GetString(template.Apply());

            var answer = new System.Text.StringBuilder();
            var watch = Stopwatch.StartNew();
            // 8 rather than 1: the model often answers "요청입니다" or leads with a particle, and
            // ParseKind reads the first of the three words out of whatever comes back. Stopping at
            // one token would turn those into 판정불가.
            //
            // Temperature 0 unlike the draft call, which is left at llama.cpp's default. Sampling is
            // what a draft needs and what a classification must not have: the first measurement ran
            // the same sentence 20 times and got 요청 ten times, 약속 once and nothing readable nine
            // times. There is one right answer here, so the model should give the same one every time
            // — and it makes the label reproducible enough to actually measure.
            var inferenceParams = new InferenceParams
            {
                MaxTokens = 8,
                SamplingPipeline = new DefaultSamplingPipeline { Temperature = 0 },
                AntiPrompts = _markers == null ? new List<string>() : new List<string> { _markers.TurnEnd },
            };

            await foreach (var token in _kindExecutor.InferAsync(prompt, inferenceParams, cancellationToken))
            {
                answer.Append(token);
            }

            _lastClassify = $"{watch.Elapsed.TotalSeconds:F1}s";
            _lastRawKind = answer.ToString().Trim();
            return DraftPromptBuilder.ParseKind(_lastRawKind);
        }
        catch (Exception)
        {
            // The draft is the product; a classifier that throws must not take it down with it.
            return null;
        }
    }

    /// <summary>
    /// The classifier on its own, with no draft behind it — one 1.4s call instead of a 15s turn.
    /// This is what makes the classifier measurable as a thing in itself: whether the label helps
    /// and whether the label is right are two different questions, and the first measurement could
    /// not tell them apart. Returns the kind, and what the model actually said, so an unreadable
    /// answer can be looked at rather than just counted as 판정불가.
    /// </summary>
    public async Task<(string? Kind, string Raw, double Seconds)> ClassifyOnlyAsync(
        string intent, CancellationToken cancellationToken = default)
    {
        await EnsureWeightsAsync(cancellationToken);
        Progress?.Report(string.Empty);
        await _inferLock.WaitAsync(cancellationToken);
        try
        {
            var watch = Stopwatch.StartNew();
            var kind = await ClassifyAsync(intent, cancellationToken);
            return (kind, _lastRawKind ?? "", watch.Elapsed.TotalSeconds);
        }
        finally
        {
            _inferLock.Release();
        }
    }

    /// <summary>
    /// Sampling is left at llama.cpp's default unless EMAILWRITE_TEMP says otherwise, so the
    /// measured baseline stays reproducible and the temperature can be A/B'd without a rebuild
    /// (same idea as EMAILWRITE_MODEL and EMAILWRITE_THREADS). Returning null means "don't set a
    /// pipeline", which is what the code did before this knob existed.
    ///
    /// Worth trying because the draft failures look like sampling rather than capability: twenty
    /// runs of one unchanged prompt produced twenty different messages, thirteen of them fine.
    /// Writing a draft from points the user already supplied does not need much creativity.
    /// </summary>
    private static float? ReadTemperature()
    {
        var raw = Environment.GetEnvironmentVariable("EMAILWRITE_TEMP");
        return float.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var temperature)
            ? temperature
            : null;
    }

    /// <summary>
    /// Cuts the reply at the first control string the model emitted. StatelessExecutor swallowed
    /// the anti-prompt it stopped on; InteractiveExecutor streams it through, so without this the
    /// draft shown to the user can end in a literal turn-end marker. Only the display copy is
    /// trimmed — the raw text is what the KV cache holds, and the next turn's continuation has to
    /// account for it.
    ///
    /// The markers come from the loaded model rather than a hardcoded family: this was ChatML for
    /// as long as the prototype ran Qwen, and kept being ChatML after the switch to Mi:dm, which
    /// renders Llama-3 headers and has no "&lt;|im_start|&gt;" in its vocabulary at all.
    /// </summary>
    internal static string StripTurnMarkers(string raw, TurnMarkers? markers)
    {
        if (markers == null)
        {
            return raw;
        }

        foreach (var marker in new[] { markers.TurnEnd, markers.UserOpen, markers.AssistantOpen })
        {
            // A marker that came back empty would make IndexOf return 0 and take the whole reply
            // with it.
            if (marker.Length == 0)
            {
                continue;
            }

            var index = raw.IndexOf(marker, StringComparison.Ordinal);
            if (index >= 0)
            {
                raw = raw[..index];
            }
        }

        return raw;
    }

    /// <summary>
    /// Backstop for the "/no_think" switch: Qwen3 still emits an (empty) &lt;think&gt; block,
    /// and ignores the switch outright on some turns. Either way the user must never see it.
    /// </summary>
    private static string StripThinking(string raw)
    {
        var end = raw.LastIndexOf("</think>", StringComparison.Ordinal);
        return (end >= 0 ? raw[(end + "</think>".Length)..] : raw).Trim();
    }

    private async Task EnsureWeightsAsync(CancellationToken cancellationToken)
    {
        if (_weights != null)
        {
            return;
        }

        await _initLock.WaitAsync(cancellationToken);
        try
        {
            if (_weights != null)
            {
                return;
            }

            if (!File.Exists(_modelPath))
            {
                // Never auto-download over an override: ModelDownloadUrl points at the default
                // model, so fetching it into the overridden filename would silently load the
                // wrong weights under the right name.
                if (!_modelPath.EndsWith(ModelFileName, StringComparison.OrdinalIgnoreCase))
                {
                    throw new FileNotFoundException(
                        $"EMAILWRITE_MODEL이 가리키는 파일이 없어요: {_modelPath}", _modelPath);
                }

                await DownloadModelAsync(cancellationToken);
            }

            // LLamaSharp defaults to half the logical processors, which on this box (8 cores, no
            // hyperthreading) left four of them idle: 18.1s -> 13.6s per turn just from asking for
            // all of them. ProcessorCount is the right default on hybrid/non-HT CPUs and merely
            // 2x-oversubscribed on hyperthreaded ones; EMAILWRITE_THREADS is the calibration knob
            // for machines where that costs more than it gains.
            var threads = int.TryParse(Environment.GetEnvironmentVariable("EMAILWRITE_THREADS"), out var t)
                ? t
                : Environment.ProcessorCount;
            var modelParams = new ModelParams(_modelPath)
            {
                ContextSize = 4096,
                Threads = threads,
                BatchThreads = threads,
                // Everything the GPU will take. On this box that is 프리필 29.9s -> 9.0s and
                // 디코드 10.2 -> 27.1 t/s (Experiments.md §D); with no Vulkan device answering,
                // llama.cpp leaves every layer on the CPU and this costs nothing. 99 rather than
                // the real layer count because it only has to be "at least all of them".
                // EMAILWRITE_GPU_LAYERS=0 forces the old CPU-only path back, which is the knob to
                // reach for if a machine turns out to have a Vulkan device slower than its CPU.
                GpuLayerCount = int.TryParse(Environment.GetEnvironmentVariable("EMAILWRITE_GPU_LAYERS"), out var gpu)
                    ? gpu
                    : 99,
            };
            _modelParams = modelParams;
            // Same weights, same threads, same GPU layers — only the context is small, because the
            // classifier prompt is one sentence and a 4096-token KV allocation per call is pure waste.
            _kindParams = new ModelParams(_modelPath)
            {
                ContextSize = 512,
                Threads = modelParams.Threads,
                BatchThreads = modelParams.BatchThreads,
                GpuLayerCount = modelParams.GpuLayerCount,
            };
            // Reading 1.06GB off disk is several seconds on its own, and on a first run it lands
            // right after the download — so the line has to change, or it reads as a stuck 100%.
            Progress?.Report("모델 불러오는 중…");
            _weights = await LLamaWeights.LoadFromFileAsync(modelParams, cancellationToken);
            _kernel = DescribeKernel();
            _markers = ProbeMarkers();
        }
        finally
        {
            _initLock.Release();
        }
    }

    /// <summary>
    /// Which ggml-cpu.dll the process ended up with, as folder name + size — nothing reported this
    /// before, and that is the only reason the app ran on the noavx kernel unnoticed for weeks
    /// (prefill 422s against 50s on avx2). <see cref="Avx2"/>.IsSupported cannot stand in for it:
    /// it answers what the CPU can do, not which build got loaded. The folder name says it when the
    /// backend layout survived the build; the size says it when it did not, because a flattened
    /// output leaves the file sitting next to the exe under no telling name
    /// (noavx 427K / avx 508K / avx2 544K / avx512 620K).
    /// </summary>
    private static string DescribeKernel()
    {
        string? cpu = null;
        var gpu = false;
        foreach (ProcessModule module in Process.GetCurrentProcess().Modules)
        {
            if (string.Equals(module.ModuleName, "ggml-cpu.dll", StringComparison.OrdinalIgnoreCase))
            {
                var folder = Path.GetFileName(Path.GetDirectoryName(module.FileName)) ?? "?";
                cpu = $"{folder} {new FileInfo(module.FileName).Length / 1024}K";
            }
            // The same question one layer up: the CPU kernel is loaded either way, so on its own
            // that name says nothing about whether the GPU took the layers. Without this the
            // status line reads identically at 8.9s and 13.8s of prefill (§D), which is exactly
            // the shape §8 took — a backend silently not being the one that was measured.
            else if (string.Equals(module.ModuleName, "ggml-vulkan.dll", StringComparison.OrdinalIgnoreCase))
            {
                gpu = true;
            }
        }

        return (cpu ?? "확인 불가") + (gpu ? " +vulkan" : "");
    }

    /// <summary>
    /// How this model opens a user turn, opens an assistant turn, and ends one. Every string is
    /// taken verbatim out of the model's own rendered template, whitespace included, so appending
    /// them reproduces exactly what <see cref="LLamaTemplate"/> would have produced.
    /// </summary>
    internal sealed record TurnMarkers(string UserOpen, string AssistantOpen, string TurnEnd);

    // Control characters, so they cannot collide with anything a chat template emits around them.
    internal const string ProbeUser = "\u0001PU\u0001";
    internal const string ProbeAssistant = "\u0002PA\u0002";
    internal const string ProbeUser2 = "\u0003PU2\u0003";

    private string Render(bool addAssistant, params (string Role, string Text)[] messages)
    {
        var template = new LLamaTemplate(_weights!) { AddAssistant = addAssistant };
        foreach (var (role, text) in messages)
        {
            template.Add(role, text);
        }

        return System.Text.Encoding.UTF8.GetString(template.Apply());
    }

    /// <summary>
    /// Renders three throwaway conversations through the model's template and reads the turn
    /// delimiters back out of them. Three rather than one because a single render cannot say where
    /// the preamble (bos token, date line, default system block) ends and the first user header
    /// begins — only a pair that differs by exactly one turn can.
    ///
    /// Deliberately never renders two user turns in a row to get there: some templates reject a
    /// conversation that does not alternate, and a probe that throws on those models would take
    /// the whole generator down at load time.
    /// </summary>
    private TurnMarkers? ProbeMarkers()
    {
        try
        {
            return ParseMarkers(
                Render(false, ("user", ProbeUser)),
                Render(true, ("user", ProbeUser)),
                Render(true, ("user", ProbeUser), ("assistant", ProbeAssistant), ("user", ProbeUser2)));
        }
        catch (Exception)
        {
            // Continuation is an optimisation; losing it costs a full prefill per turn, which is
            // what this class did before the KV cache existed. Feeding the model markers invented
            // from a failed parse is the thing actually worth avoiding.
            return null;
        }
    }

    /// <summary>
    /// Pure string work, kept separate from <see cref="ProbeMarkers"/> so it can be checked without
    /// loading 1.3GB of weights. Given
    /// <paramref name="userOnly"/>   = preamble + userOpen + PU + turnEnd,
    /// <paramref name="userThenGen"/> = the same + assistantOpen, and
    /// <paramref name="fullRound"/>  = ... + PA + turnEnd + userOpen + PU2 + ...,
    /// each marker falls out by subtraction.
    /// </summary>
    internal static TurnMarkers? ParseMarkers(string userOnly, string userThenGen, string fullRound)
    {
        var turnEnd = After(userOnly, ProbeUser);
        var assistantOpen = After(userThenGen, ProbeUser);
        var afterAssistant = After(fullRound, ProbeAssistant);

        if (turnEnd.Length == 0
            || !assistantOpen.StartsWith(turnEnd, StringComparison.Ordinal)
            || !afterAssistant.StartsWith(turnEnd, StringComparison.Ordinal))
        {
            return null;
        }

        assistantOpen = assistantOpen[turnEnd.Length..];
        var rest = afterAssistant[turnEnd.Length..];
        var userAt = rest.IndexOf(ProbeUser2, StringComparison.Ordinal);
        if (userAt < 0 || assistantOpen.Length == 0)
        {
            return null;
        }

        return new TurnMarkers(rest[..userAt], assistantOpen, turnEnd);
    }

    private static string After(string haystack, string needle)
    {
        var index = haystack.IndexOf(needle, StringComparison.Ordinal);
        return index < 0 ? string.Empty : haystack[(index + needle.Length)..];
    }

    /// <summary>
    /// Throws away the live KV cache and starts a fresh one. Called whenever the next prompt
    /// isn't a continuation of what the context already holds — a different mail, a reopened
    /// popup, or a turn that failed partway through.
    /// </summary>
    private void StartSession()
    {
        _executor = null;
        _context?.Dispose();
        _context = new LLamaContext(_weights!, _modelParams!);
        _executor = new InteractiveExecutor(_context);
        _sessionId = -1;
        _fedMessageCount = 0;
        _kvEndsAtTurnEnd = false;
    }

    /// <summary>
    /// The markup the cache still needs in order to continue: close the reply that is already in
    /// it (unless generation stopped on the closing token and it went in too), then the new
    /// messages, then the header that hands the turn back to the model. Assembled here rather than
    /// via <see cref="LLamaTemplate"/>, which can only render a conversation from the system prompt
    /// down — the whole point here is to render only the tail. The pieces are still the model's
    /// own, taken from <see cref="ProbeMarkers"/> rather than assumed to be ChatML.
    /// </summary>
    private string BuildContinuation(IEnumerable<ChatMessage> newMessages)
    {
        var markers = _markers!;
        var sb = new System.Text.StringBuilder();
        if (!_kvEndsAtTurnEnd)
        {
            sb.Append(markers.TurnEnd);
        }

        foreach (var message in newMessages)
        {
            sb.Append(message.Role == ChatRole.User ? markers.UserOpen : markers.AssistantOpen)
              .Append(message.Text)
              .Append(markers.TurnEnd);
        }

        return sb.Append(markers.AssistantOpen).ToString();
    }

    private async Task DownloadModelAsync(CancellationToken cancellationToken)
    {
        var tempPath = _modelPath + ".download";
        using var http = new HttpClient();
        // No timeout: HttpClient's default 100s applies to the whole response, and 1.06GB over an
        // office link does not fit in it. ResponseHeadersRead already means the body streams.
        http.Timeout = Timeout.InfiniteTimeSpan;
        using var response = await http.GetAsync(
            ModelDownloadUrl, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
        response.EnsureSuccessStatusCode();

        var total = response.Content.Headers.ContentLength;
        // Hand-rolled instead of CopyToAsync purely to count bytes. Reported every 8MB rather than
        // every chunk: ~175 updates over the whole file instead of ~18,000.
        const int ReportEvery = 8 * 1024 * 1024;
        var buffer = new byte[81920];
        long read = 0, reportedAt = 0;

        await using (var fileStream = File.Create(tempPath))
        await using (var httpStream = await response.Content.ReadAsStreamAsync(cancellationToken))
        {
            Progress?.Report(FormatDownloadProgress(0, total));
            int count;
            while ((count = await httpStream.ReadAsync(buffer, cancellationToken)) > 0)
            {
                await fileStream.WriteAsync(buffer.AsMemory(0, count), cancellationToken);
                read += count;
                if (read - reportedAt >= ReportEvery)
                {
                    reportedAt = read;
                    Progress?.Report(FormatDownloadProgress(read, total));
                }
            }
        }

        // Verified before the file gets its real name, so a bad download is never loadable and
        // never has to be diagnosed: the weights come from a third party's conversion of NAVER's
        // model, over a link that can also just truncate. Both failures look identical from here
        // and this catches both. Checked on the temp file rather than on every startup — hashing
        // 1.06GB costs several seconds and the cache does not rot on its own.
        string actual;
        await using (var downloaded = File.OpenRead(tempPath))
        {
            actual = Convert.ToHexString(
                await SHA256.HashDataAsync(downloaded, cancellationToken)).ToLowerInvariant();
        }

        if (actual != ModelSha256)
        {
            File.Delete(tempPath);
            throw new InvalidDataException(
                $"내려받은 모델 파일이 손상됐거나 바뀌었어요. 다시 시도해 주세요. "
                + $"(기대 {ModelSha256[..12]}…, 실제 {actual[..12]}…)");
        }

        File.Move(tempPath, _modelPath);
    }

    /// <summary>
    /// The download line. Hugging Face does send Content-Length, but a proxy or a chunked
    /// re-encode can drop it, and a progress line that divides by null is a worse first run than
    /// one that just counts up.
    /// </summary>
    internal static string FormatDownloadProgress(long read, long? total)
    {
        var readGb = read / (1024.0 * 1024 * 1024);
        if (total is not > 0)
        {
            return $"모델 내려받는 중… {readGb:0.0}GB";
        }

        return $"모델 내려받는 중… {read * 100 / total.Value}% ({readGb:0.0}/{total.Value / (1024.0 * 1024 * 1024):0.0}GB)";
    }

    /// <summary>
    /// Renders the system prompt and full visible conversation through the model's own
    /// chat template — whatever that model's is — mirroring <see cref="ClaudeDraftGenerator"/>'s
    /// "resend the whole history every turn" approach instead of keeping session state.
    /// </summary>
    private string BuildPrompt(EmailContext? emailContext, IReadOnlyList<ChatMessage> conversation, string? kind)
    {
        // AddAssistant appends the header this model uses to open an assistant turn — the same
        // string ProbeMarkers reads back as TurnMarkers.AssistantOpen. It defaults to false, and
        // without it the model is handed a conversation that never gives it the turn.
        // "/no_think" is Qwen3's soft switch for its reasoning mode — thinking triples latency
        // (~43s vs ~13s) for a two-sentence draft. It means nothing to other models, so it's
        // only appended for Qwen to keep a stray token out of everyone else's system prompt.
        var isQwen = _modelPath.Contains("qwen", StringComparison.OrdinalIgnoreCase);
        var (systemPrompt, examples) = DraftPromptBuilder.PromptFor(emailContext, kind);
        var template = new LLamaTemplate(_weights!) { AddAssistant = true };
        template.Add("system", systemPrompt + (isQwen ? "\n/no_think" : ""));
        foreach (var (role, text) in examples)
        {
            template.Add(role == ChatRole.User ? "user" : "assistant", text);
        }

        var (openingTurn, rest) = DraftPromptBuilder.BuildTurns(emailContext, conversation, kind);
        template.Add("user", openingTurn);

        foreach (var message in rest)
        {
            template.Add(message.Role == ChatRole.User ? "user" : "assistant", message.Text);
        }

        return System.Text.Encoding.UTF8.GetString(template.Apply());
    }

    public void Dispose()
    {
        _context?.Dispose();
        _weights?.Dispose();
        _initLock.Dispose();
        _inferLock.Dispose();
    }
}
