using LLama;
using LLama.Common;
using EmailWrite.Core.OutlookIntegration;

namespace EmailWrite.Core.Drafting;

/// <summary>
/// Runs draft generation locally via LLamaSharp against a Qwen3-1.7B GGUF, downloaded
/// on demand into a machine-wide model cache. M8's minimal slice: just enough
/// to unblock M6/M7 testing without an Anthropic API key — mode-switch UI, custom model
/// URLs, and proper minimum-spec detection are M9.
/// </summary>
public sealed class LocalDraftGenerator : IDraftGenerator, IDisposable
{
    private const string ModelFileName = "midm-2.0-mini-instruct-q4_k_m.gguf";
    private const string ModelDownloadUrl =
        "https://huggingface.co/mykor/Midm-2.0-Mini-Instruct-gguf/resolve/main/Midm-2.0-Mini-Instruct-Q4_K_M.gguf";

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
    // Which conversation the cache belongs to. -1 means "nothing usable in there". This used to
    // compare the EmailContext by reference, which worked only as long as there was always a mail:
    // with free-writing both sides are null, null equals null, and a brand new conversation would
    // look like a continuation of the last one.
    private int _sessionId = -1;
    private int _fedMessageCount;
    private bool _kvEndsAtTurnEnd;

    public LocalDraftGenerator()
    {
        // Deliberately NOT %LocalAppData%: llama.cpp's gguf_init_from_file opens the model
        // with plain ANSI fopen, so a non-ASCII user profile name (e.g. C:\Users\조하연\...)
        // arrives mangled and the load fails. CommonApplicationData is C:\ProgramData —
        // always ASCII, writable by Users, and shared so a second user on the same PC
        // doesn't re-download 1.1GB.
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
        try
        {
            await EnsureWeightsAsync(cancellationToken);
        }
        catch (Exception ex)
        {
            return $"로컬 모델을 준비하는 중 오류가 발생했어요(AVX2 미지원 CPU이거나 다운로드에 실패했을 수 있어요): {ex.Message}";
        }

        await _inferLock.WaitAsync(cancellationToken);
        try
        {
            // Continue the live cache only when this is the same conversation and it has merely
            // grown; a reopened popup arrives with a new session id and starts over. The mail is
            // fixed when the session opens, so the id alone settles both questions.
            string toFeed;
            if (_executor != null && _sessionId == sessionId && conversation.Count > _fedMessageCount)
            {
                toFeed = BuildContinuation(conversation.Skip(_fedMessageCount));
            }
            else
            {
                StartSession();
                _sessionId = sessionId;
                toFeed = BuildPrompt(emailContext, conversation);
            }

            var inferenceParams = new InferenceParams
            {
                MaxTokens = 512,
                AntiPrompts = new List<string> { "<|im_end|>", "<|im_start|>" },
            };

            var text = new System.Text.StringBuilder();
            await foreach (var token in _executor!.InferAsync(toFeed, inferenceParams, cancellationToken))
            {
                text.Append(token);
            }

            // The cache now also holds the reply just generated, which the caller will hand back
            // as the next conversation entry — hence +1.
            var raw = text.ToString();
            _fedMessageCount = conversation.Count + 1;
            _kvEndsAtTurnEnd = raw.TrimEnd().EndsWith("<|im_end|>", StringComparison.Ordinal);

            var result = StripThinking(StripChatMarkers(raw));
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
    /// Cuts the reply at the first ChatML control token. StatelessExecutor swallowed the
    /// anti-prompt it stopped on; InteractiveExecutor streams it through, so without this the
    /// draft shown to the user ends in a literal "&lt;|im_end|&gt;". Only the display copy is
    /// trimmed — the raw text stays in _fedPrompt, because those tokens really are in the KV
    /// cache and the next turn's prefix has to account for them.
    /// </summary>
    internal static string StripChatMarkers(string raw)
    {
        foreach (var marker in new[] { "<|im_end|>", "<|im_start|>" })
        {
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
            };
            _modelParams = modelParams;
            _weights = await LLamaWeights.LoadFromFileAsync(modelParams, cancellationToken);
        }
        finally
        {
            _initLock.Release();
        }
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
    /// The ChatML the cache still needs in order to continue: close the reply that is already in
    /// it (unless generation stopped on the closing token and it went in too), then the new
    /// messages, then the header that hands the turn back to the model. Written out by hand
    /// rather than via <see cref="LLamaTemplate"/>, which can only render a conversation from the
    /// system prompt down — the whole point here is to render only the tail.
    /// </summary>
    private string BuildContinuation(IEnumerable<ChatMessage> newMessages)
    {
        var sb = new System.Text.StringBuilder();
        if (!_kvEndsAtTurnEnd)
        {
            sb.Append("<|im_end|>");
        }

        sb.Append('\n');
        foreach (var message in newMessages)
        {
            sb.Append("<|im_start|>")
              .Append(message.Role == ChatRole.User ? "user" : "assistant")
              .Append('\n')
              .Append(message.Text)
              .Append("<|im_end|>\n");
        }

        return sb.Append("<|im_start|>assistant\n").ToString();
    }

    private async Task DownloadModelAsync(CancellationToken cancellationToken)
    {
        var tempPath = _modelPath + ".download";
        using var http = new HttpClient();
        using var response = await http.GetAsync(
            ModelDownloadUrl, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
        response.EnsureSuccessStatusCode();

        await using (var fileStream = File.Create(tempPath))
        await using (var httpStream = await response.Content.ReadAsStreamAsync(cancellationToken))
        {
            await httpStream.CopyToAsync(fileStream, cancellationToken);
        }

        File.Move(tempPath, _modelPath);
    }

    /// <summary>
    /// Renders the system prompt and full visible conversation through the model's own
    /// chat template (Qwen2.5's ChatML), mirroring <see cref="ClaudeDraftGenerator"/>'s
    /// "resend the whole history every turn" approach instead of keeping session state.
    /// </summary>
    private string BuildPrompt(EmailContext? emailContext, IReadOnlyList<ChatMessage> conversation)
    {
        // AddAssistant appends the "<|im_start|>assistant" header the model needs to start
        // its turn. It defaults to false, and without it generation trips the "<|im_start|>"
        // anti-prompt on the very first token and returns nothing.
        // "/no_think" is Qwen3's soft switch for its reasoning mode — thinking triples latency
        // (~43s vs ~13s) for a two-sentence draft. It means nothing to other models, so it's
        // only appended for Qwen to keep a stray token out of everyone else's system prompt.
        var isQwen = _modelPath.Contains("qwen", StringComparison.OrdinalIgnoreCase);
        var (systemPrompt, examples) = DraftPromptBuilder.PromptFor(emailContext);
        var template = new LLamaTemplate(_weights!) { AddAssistant = true };
        template.Add("system", systemPrompt + (isQwen ? "\n/no_think" : ""));
        foreach (var (role, text) in examples)
        {
            template.Add(role == ChatRole.User ? "user" : "assistant", text);
        }

        var (openingTurn, rest) = DraftPromptBuilder.BuildTurns(emailContext, conversation);
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
