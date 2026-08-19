using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace EmailWrite.Core.Drafting;

/// <summary>
/// Where an API key typed into the popup is kept, so a colleague who installs the app can paste
/// theirs once instead of setting a machine environment variable and restarting.
///
/// The env variable is still read as a fallback, and only as a fallback: the typed key has to win,
/// otherwise entering one on a PC that already has a stale variable (which is exactly how
/// ANTHROPIC_API_KEY died here on 2026-08-07) would silently keep using the dead one.
/// </summary>
public static class ApiKeys
{
    // LocalApplicationData, not Roaming: a DPAPI blob is bound to this user *on this machine*, so a
    // roaming profile would carry a file to another PC that can never be decrypted there — a
    // "저장했는데 안 먹는" bug with nothing on screen to explain it. Non-ASCII path is fine here;
    // that only breaks llama.cpp's ANSI fopen, not .NET's file APIs.
    private static readonly string FilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "EmailWrite", "keys.json");

    /// <summary>
    /// The key under <paramref name="name"/>, or null when neither the file nor the environment
    /// has one. <paramref name="name"/> is the environment variable's name (UPSTAGE_API_KEY) so
    /// there is one vocabulary rather than two.
    /// </summary>
    public static string? Get(string name)
    {
        if (Load().TryGetValue(name, out var stored) && stored.Length > 0)
        {
            try
            {
                return Encoding.UTF8.GetString(ProtectedData.Unprotect(
                    Convert.FromBase64String(stored), null, DataProtectionScope.CurrentUser));
            }
            catch (Exception ex) when (ex is CryptographicException or FormatException)
            {
                // Copied in from another account, or half-written. Falling through to the
                // environment beats throwing at the one moment the user asked for a draft.
            }
        }

        var env = Environment.GetEnvironmentVariable(name);
        return string.IsNullOrWhiteSpace(env) ? null : env;
    }

    /// <summary>
    /// Stores <paramref name="value"/> under <paramref name="name"/>, encrypted with DPAPI so the
    /// file is worthless if it is copied off the PC — an API key sitting in cleartext under the
    /// user profile ends up in backups and support zips. An empty value removes the entry, which
    /// is both "키 지우기" and what lets the self-check run without leaving a fake key behind.
    /// </summary>
    public static void Set(string name, string value)
    {
        var all = Load();
        if (string.IsNullOrWhiteSpace(value))
        {
            all.Remove(name);
        }
        else
        {
            all[name] = Convert.ToBase64String(ProtectedData.Protect(
                Encoding.UTF8.GetBytes(value), null, DataProtectionScope.CurrentUser));
        }

        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        File.WriteAllText(FilePath, JsonSerializer.Serialize(all));
    }

    private static Dictionary<string, string> Load()
    {
        try
        {
            return JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(FilePath))
                ?? new Dictionary<string, string>();
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            // No file yet is the normal first run, and a corrupt one has to behave the same way:
            // the next Set rewrites it whole.
            return new Dictionary<string, string>();
        }
    }
}
