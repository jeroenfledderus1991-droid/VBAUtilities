using System;
using System.IO;
using System.Net;

namespace VBEAddIn
{
    internal static class GitHubReleaseChecker
    {
        private const string LatestReleaseApiUrl = "https://api.github.com/repos/jeroenfledderus1991-droid/VBAUtilities/releases/latest";

        internal static bool TryGetLatestRelease(out string version, out string releaseUrl, out string failureReason)
        {
            version = string.Empty;
            releaseUrl = string.Empty;
            failureReason = string.Empty;

            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(LatestReleaseApiUrl);
                request.Method = "GET";
                request.Accept = "application/vnd.github+json";
                request.UserAgent = "VBEAddIn/1.0";
                request.Timeout = 2500;
                request.ReadWriteTimeout = 2500;

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    string json = reader.ReadToEnd();
                    string tag = ExtractJsonStringValue(json, "tag_name");
                    releaseUrl = ExtractJsonStringValue(json, "html_url");

                    if (string.IsNullOrWhiteSpace(tag))
                    {
                        failureReason = "Geen tag_name ontvangen van GitHub.";
                        return false;
                    }

                    version = NormalizeTagToVersion(tag);
                    if (string.IsNullOrWhiteSpace(version))
                    {
                        failureReason = "Kon tag niet omzetten naar versie: " + tag;
                        return false;
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                failureReason = ex.Message;
                return false;
            }
        }

        internal static bool IsRemoteNewer(string currentVersion, string remoteVersion)
        {
            Version current = ParseVersion(currentVersion);
            Version remote = ParseVersion(remoteVersion);

            if (current == null || remote == null)
            {
                return false;
            }

            return remote > current;
        }

        private static Version ParseVersion(string value)
        {
            Version parsed;
            if (Version.TryParse(value ?? string.Empty, out parsed))
            {
                return parsed;
            }

            return null;
        }

        private static string NormalizeTagToVersion(string tag)
        {
            if (string.IsNullOrWhiteSpace(tag))
            {
                return string.Empty;
            }

            string trimmed = tag.Trim();
            if (trimmed.StartsWith("v", StringComparison.OrdinalIgnoreCase))
            {
                trimmed = trimmed.Substring(1);
            }

            return trimmed.Trim();
        }

        private static string ExtractJsonStringValue(string json, string propertyName)
        {
            if (string.IsNullOrEmpty(json) || string.IsNullOrEmpty(propertyName))
            {
                return string.Empty;
            }

            string key = "\"" + propertyName + "\"";
            int keyIndex = json.IndexOf(key, StringComparison.Ordinal);
            if (keyIndex < 0)
            {
                return string.Empty;
            }

            int colonIndex = json.IndexOf(':', keyIndex + key.Length);
            if (colonIndex < 0)
            {
                return string.Empty;
            }

            int firstQuote = json.IndexOf('"', colonIndex + 1);
            if (firstQuote < 0)
            {
                return string.Empty;
            }

            int secondQuote = json.IndexOf('"', firstQuote + 1);
            while (secondQuote > 0 && json[secondQuote - 1] == '\\')
            {
                secondQuote = json.IndexOf('"', secondQuote + 1);
            }

            if (secondQuote < 0)
            {
                return string.Empty;
            }

            return json.Substring(firstQuote + 1, secondQuote - firstQuote - 1)
                .Replace("\\/", "/")
                .Replace("\\\"", "\"");
        }
    }
}
