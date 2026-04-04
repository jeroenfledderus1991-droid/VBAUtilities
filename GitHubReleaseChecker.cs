using System;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;

namespace VBEAddIn
{
    internal static class GitHubReleaseChecker
    {
        private const string RepoOwner = "jeroenfledderus1991-droid";
        private const string RepoName = "VBAUtilities";
        private const string InstallerFileName = "VBEAddIn-Installer.exe";
        private const string LatestReleaseApiUrl = "https://api.github.com/repos/jeroenfledderus1991-droid/VBAUtilities/releases/latest";
        private const string LatestTagApiUrl = "https://api.github.com/repos/jeroenfledderus1991-droid/VBAUtilities/tags?per_page=1";
        private const string LatestInstallerDownloadUrl = "https://github.com/jeroenfledderus1991-droid/VBAUtilities/releases/latest/download/VBEAddIn-Installer.exe";

        internal static bool TryGetLatestRelease(out string version, out string releaseUrl, out string installerUrl, out string failureReason)
        {
            version = string.Empty;
            releaseUrl = string.Empty;
            installerUrl = string.Empty;
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
                    installerUrl = ExtractPreferredInstallerUrl(json);

                    if (string.IsNullOrWhiteSpace(installerUrl))
                    {
                        installerUrl = BuildInstallerDownloadUrl(tag);
                    }

                    if (string.IsNullOrWhiteSpace(installerUrl))
                    {
                        installerUrl = LatestInstallerDownloadUrl;
                    }

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
            catch (WebException ex)
            {
                HttpWebResponse response = ex.Response as HttpWebResponse;
                if (response != null && response.StatusCode == HttpStatusCode.NotFound)
                {
                    // Fallback: als er nog geen GitHub Release bestaat, gebruik de nieuwste tag.
                    return TryGetLatestTag(out version, out releaseUrl, out installerUrl, out failureReason);
                }

                failureReason = ex.Message;
                return false;
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

        private static bool TryGetLatestTag(out string version, out string releaseUrl, out string installerUrl, out string failureReason)
        {
            version = string.Empty;
            releaseUrl = string.Empty;
            installerUrl = string.Empty;
            failureReason = string.Empty;

            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(LatestTagApiUrl);
                request.Method = "GET";
                request.Accept = "application/vnd.github+json";
                request.UserAgent = "VBEAddIn/1.0";
                request.Timeout = 2500;
                request.ReadWriteTimeout = 2500;

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    string json = reader.ReadToEnd();
                    Match match = Regex.Match(json, "\\\"name\\\"\\s*:\\s*\\\"([^\\\"]+)\\\"", RegexOptions.IgnoreCase);
                    if (!match.Success || match.Groups.Count < 2)
                    {
                        failureReason = "Geen tag gevonden via GitHub tags API.";
                        return false;
                    }

                    string tag = match.Groups[1].Value;
                    version = NormalizeTagToVersion(tag);
                    releaseUrl = "https://github.com/jeroenfledderus1991-droid/VBAUtilities/tree/" + tag;
                    installerUrl = BuildInstallerDownloadUrl(tag);

                    if (string.IsNullOrWhiteSpace(installerUrl))
                    {
                        installerUrl = LatestInstallerDownloadUrl;
                    }

                    return !string.IsNullOrWhiteSpace(version);
                }
            }
            catch (Exception ex)
            {
                failureReason = ex.Message;
                return false;
            }
        }

        private static string ExtractPreferredInstallerUrl(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return string.Empty;
            }

            string firstExeOrMsi = string.Empty;
            MatchCollection matches = Regex.Matches(
                json,
                "\\\"browser_download_url\\\"\\s*:\\s*\\\"([^\\\"]+)\\\"",
                RegexOptions.IgnoreCase);

            foreach (Match match in matches)
            {
                if (match.Groups.Count < 2)
                {
                    continue;
                }

                string url = match.Groups[1].Value
                    .Replace("\\/", "/")
                    .Replace("\\\"", "\"");

                string lower = url.ToLowerInvariant();
                bool isInstallerFile = lower.EndsWith(".exe") || lower.EndsWith(".msi");

                if (isInstallerFile && string.IsNullOrWhiteSpace(firstExeOrMsi))
                {
                    firstExeOrMsi = url;
                }

                if (isInstallerFile && lower.Contains("installer"))
                {
                    return url;
                }
            }

            return firstExeOrMsi;
        }

        private static string BuildInstallerDownloadUrl(string tag)
        {
            string normalizedTag = (tag ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalizedTag))
            {
                return string.Empty;
            }

            return "https://github.com/" + RepoOwner + "/" + RepoName + "/releases/download/" + normalizedTag + "/" + InstallerFileName;
        }
    }
}
