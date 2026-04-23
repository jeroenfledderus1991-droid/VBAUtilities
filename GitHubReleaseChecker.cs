using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace VBEAddIn
{
    [DataContract]
    internal sealed class GitHubAssetDto
    {
        [DataMember(Name = "name")]
        public string Name { get; set; }

        [DataMember(Name = "browser_download_url")]
        public string BrowserDownloadUrl { get; set; }
    }

    [DataContract]
    internal sealed class GitHubReleaseDto
    {
        [DataMember(Name = "tag_name")]
        public string TagName { get; set; }

        [DataMember(Name = "html_url")]
        public string HtmlUrl { get; set; }

        [DataMember(Name = "name")]
        public string Name { get; set; }

        [DataMember(Name = "body")]
        public string Body { get; set; }

        [DataMember(Name = "published_at")]
        public string PublishedAt { get; set; }

        [DataMember(Name = "draft")]
        public bool Draft { get; set; }

        [DataMember(Name = "prerelease")]
        public bool Prerelease { get; set; }

        [DataMember(Name = "assets")]
        public List<GitHubAssetDto> Assets { get; set; }
    }

    internal sealed class GitHubReleaseInfo
    {
        public string Version { get; set; }
        public string Tag { get; set; }
        public string ReleaseUrl { get; set; }
        public string InstallerUrl { get; set; }
        public string InstallerFileName { get; set; }
        public string Body { get; set; }
        public DateTime? PublishedAtUtc { get; set; }

        public string PublishedDisplay
        {
            get
            {
                if (!PublishedAtUtc.HasValue)
                {
                    return "onbekend";
                }

                return PublishedAtUtc.Value.ToLocalTime().ToString("yyyy-MM-dd");
            }
        }

        public override string ToString()
        {
            return string.Format("{0} ({1})", Version, PublishedDisplay);
        }
    }

    internal static class GitHubReleaseChecker
    {
        private const string RepoOwner = "jeroenfledderus1991-droid";
        private const string RepoName = "VBAUtilities";
        private const string InstallerFileName = "VBEAddIn-Installer.exe";
        private const string ReleasesApiUrl = "https://api.github.com/repos/jeroenfledderus1991-droid/VBAUtilities/releases?per_page=20";
        private const string LatestTagApiUrl = "https://api.github.com/repos/jeroenfledderus1991-droid/VBAUtilities/tags?per_page=1";
        private const string LatestInstallerDownloadUrl = "https://github.com/jeroenfledderus1991-droid/VBAUtilities/releases/latest/download/VBEAddIn-Installer.exe";

        internal static bool TryGetAvailableReleases(out List<GitHubReleaseInfo> releases, out string failureReason)
        {
            releases = new List<GitHubReleaseInfo>();
            failureReason = string.Empty;

            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(ReleasesApiUrl);
                request.Method = "GET";
                request.Accept = "application/vnd.github+json";
                request.UserAgent = "VBEAddIn/1.0";
                request.Timeout = 4000;
                request.ReadWriteTimeout = 4000;

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (Stream responseStream = response.GetResponseStream())
                {
                    if (responseStream == null)
                    {
                        failureReason = "GitHub stuurde geen response stream terug.";
                        return false;
                    }

                    var serializer = new DataContractJsonSerializer(typeof(List<GitHubReleaseDto>));
                    var dtoList = serializer.ReadObject(responseStream) as List<GitHubReleaseDto>;
                    if (dtoList == null)
                    {
                        failureReason = "GitHub releases konden niet worden gelezen.";
                        return false;
                    }

                    releases = dtoList
                        .Where(r => r != null && !r.Draft && !r.Prerelease)
                        .Select(ToReleaseInfo)
                        .Where(r => !string.IsNullOrWhiteSpace(r.Version))
                        .OrderByDescending(r => ParseVersion(r.Version))
                        .ToList();

                    if (releases.Count == 0)
                    {
                        failureReason = "Geen publieke releases gevonden op GitHub.";
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
                    GitHubReleaseInfo fallback;
                    if (TryGetLatestTagRelease(out fallback, out failureReason))
                    {
                        releases.Add(fallback);
                        return true;
                    }

                    return false;
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

        internal static bool TryGetLatestRelease(out string version, out string releaseUrl, out string installerUrl, out string failureReason)
        {
            version = string.Empty;
            releaseUrl = string.Empty;
            installerUrl = string.Empty;
            failureReason = string.Empty;

            List<GitHubReleaseInfo> releases;
            if (!TryGetAvailableReleases(out releases, out failureReason))
            {
                return false;
            }

            GitHubReleaseInfo latest = releases.FirstOrDefault();
            if (latest == null)
            {
                failureReason = "Geen release-informatie beschikbaar.";
                return false;
            }

            version = latest.Version;
            releaseUrl = latest.ReleaseUrl;
            installerUrl = latest.InstallerUrl;
            return true;
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

        private static GitHubReleaseInfo ToReleaseInfo(GitHubReleaseDto dto)
        {
            string tag = dto.TagName ?? string.Empty;
            string version = NormalizeTagToVersion(tag);
            string installerUrl = ExtractInstallerUrl(dto);

            DateTime published;
            DateTime? publishedUtc = DateTime.TryParse(dto.PublishedAt, out published)
                ? published.ToUniversalTime()
                : (DateTime?)null;

            return new GitHubReleaseInfo
            {
                Version = version,
                Tag = tag,
                ReleaseUrl = dto.HtmlUrl ?? string.Empty,
                InstallerUrl = installerUrl,
                InstallerFileName = InstallerFileName,
                Body = dto.Body ?? string.Empty,
                PublishedAtUtc = publishedUtc
            };
        }

        private static string ExtractInstallerUrl(GitHubReleaseDto dto)
        {
            if (dto.Assets != null)
            {
                GitHubAssetDto namedInstaller = dto.Assets.FirstOrDefault(asset =>
                    asset != null &&
                    string.Equals(asset.Name, InstallerFileName, StringComparison.OrdinalIgnoreCase));

                if (namedInstaller != null && !string.IsNullOrWhiteSpace(namedInstaller.BrowserDownloadUrl))
                {
                    return namedInstaller.BrowserDownloadUrl;
                }

                GitHubAssetDto anyInstaller = dto.Assets.FirstOrDefault(asset =>
                    asset != null &&
                    !string.IsNullOrWhiteSpace(asset.Name) &&
                    asset.Name.IndexOf("installer", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    !string.IsNullOrWhiteSpace(asset.BrowserDownloadUrl));

                if (anyInstaller != null)
                {
                    return anyInstaller.BrowserDownloadUrl;
                }
            }

            return BuildInstallerDownloadUrl(dto.TagName);
        }

        private static bool TryGetLatestTagRelease(out GitHubReleaseInfo release, out string failureReason)
        {
            release = null;
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
                using (Stream responseStream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(responseStream))
                {
                    string json = reader.ReadToEnd();
                    int nameStart = json.IndexOf("\"name\"", StringComparison.OrdinalIgnoreCase);
                    if (nameStart < 0)
                    {
                        failureReason = "Geen tag gevonden via GitHub tags API.";
                        return false;
                    }

                    int colon = json.IndexOf(':', nameStart);
                    int firstQuote = json.IndexOf('"', colon + 1);
                    int secondQuote = json.IndexOf('"', firstQuote + 1);
                    string tag = json.Substring(firstQuote + 1, secondQuote - firstQuote - 1);
                    string version = NormalizeTagToVersion(tag);

                    release = new GitHubReleaseInfo
                    {
                        Version = version,
                        Tag = tag,
                        ReleaseUrl = "https://github.com/" + RepoOwner + "/" + RepoName + "/tree/" + tag,
                        InstallerUrl = BuildInstallerDownloadUrl(tag),
                        InstallerFileName = InstallerFileName,
                        Body = string.Empty,
                        PublishedAtUtc = null
                    };

                    return true;
                }
            }
            catch (Exception ex)
            {
                failureReason = ex.Message;
                return false;
            }
        }

        private static string BuildInstallerDownloadUrl(string tag)
        {
            string normalizedTag = (tag ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalizedTag))
            {
                return LatestInstallerDownloadUrl;
            }

            return "https://github.com/" + RepoOwner + "/" + RepoName + "/releases/download/" + normalizedTag + "/" + InstallerFileName;
        }

        private static Version ParseVersion(string value)
        {
            Version parsed;
            if (Version.TryParse(value ?? string.Empty, out parsed))
            {
                return parsed;
            }

            return new Version(0, 0, 0, 0);
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
    }
}
