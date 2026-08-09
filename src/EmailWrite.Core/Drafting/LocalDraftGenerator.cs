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
    private StatelessExecutor? _executor;

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
        CancellationToken cancellationToken = default)
    {
        if (emailContext == null)
        {
            return "Outlook에서 열려있거나 선택된 메일을 찾지 못했어요. 메일을 연 뒤 다시 열어주세요.";
        }

        StatelessExecutor executor;
        try
        {
            executor = await EnsureReadyAsync(cancellationToken);
        }
        catch (Exception ex)
        {
            return $"로컬 모델을 준비하는 중 오류가 발생했어요(AVX2 미지원 CPU이거나 다운로드에 실패했을 수 있어요): {ex.Message}";
        }

        var prompt = BuildPrompt(emailContext, conversation);

        await _inferLock.WaitAsync(cancellationToken);
        try
        {
            var inferenceParams = new InferenceParams
            {
                MaxTokens = 512,
                AntiPrompts = new List<string> { "<|im_end|>", "<|im_start|>" },
            };

            var text = new System.Text.StringBuilder();
            await foreach (var token in executor.InferAsync(prompt, inferenceParams, cancellationToken))
            {
                text.Append(token);
            }

            var result = StripThinking(text.ToString());
            return string.IsNullOrWhiteSpace(result) ? "응답을 받지 못했어요." : result;
        }
        catch (Exception ex)
        {
            return $"로컬 모델 추론 중 오류가 발생했어요: {ex.Message}";
        }
        finally
        {
            _inferLock.Release();
        }
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

    private async Task<StatelessExecutor> EnsureReadyAsync(CancellationToken cancellationToken)
    {
        if (_executor != null)
        {
            return _executor;
        }

        await _initLock.WaitAsync(cancellationToken);
        try
        {
            if (_executor != null)
            {
                return _executor;
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
            _weights = await LLamaWeights.LoadFromFileAsync(modelParams, cancellationToken);
            _executor = new StatelessExecutor(_weights, modelParams);
            return _executor;
        }
        finally
        {
            _initLock.Release();
        }
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
    private string BuildPrompt(EmailContext emailContext, IReadOnlyList<ChatMessage> conversation)
    {
        // AddAssistant appends the "<|im_start|>assistant" header the model needs to start
        // its turn. It defaults to false, and without it generation trips the "<|im_start|>"
        // anti-prompt on the very first token and returns nothing.
        // "/no_think" is Qwen3's soft switch for its reasoning mode — thinking triples latency
        // (~43s vs ~13s) for a two-sentence draft. It means nothing to other models, so it's
        // only appended for Qwen to keep a stray token out of everyone else's system prompt.
        var isQwen = _modelPath.Contains("qwen", StringComparison.OrdinalIgnoreCase);
        var template = new LLamaTemplate(_weights!) { AddAssistant = true };
        template.Add("system", DraftPromptBuilder.SystemPrompt + (isQwen ? "\n/no_think" : ""));
        template.Add("user", DraftPromptBuilder.Example.User);
        template.Add("assistant", DraftPromptBuilder.Example.Assistant);
        template.Add("user", DraftPromptBuilder.BuildContextTurn(emailContext));

        foreach (var message in conversation)
        {
            template.Add(message.Role == ChatRole.User ? "user" : "assistant", message.Text);
        }

        return System.Text.Encoding.UTF8.GetString(template.Apply());
    }

    public void Dispose()
    {
        _weights?.Dispose();
        _initLock.Dispose();
        _inferLock.Dispose();
    }
}
