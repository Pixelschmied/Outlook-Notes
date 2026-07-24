using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Reflection;
using System.Web.Script.Serialization;

namespace EmailNotes
{
    /// <summary>
    /// Best-effort update check against the project's latest GitHub release.
    /// Purely optional and non-blocking — if there is no internet or the request
    /// fails, it silently returns null and the offline notepad is unaffected.
    /// </summary>
    internal static class Updater
    {
        private const string ApiUrl =
            "https://api.github.com/repos/Pixelschmied/Outlook-Notes/releases/latest";

        public class Info
        {
            public string Version;
            public string DownloadUrl;
            public string PageUrl;
        }

        /// <summary>Returns update info if a newer release exists, else null. Safe to call off the UI thread.</summary>
        public static Info Check()
        {
            try
            {
                try { ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12; } catch { }
                var req = (HttpWebRequest)WebRequest.Create(ApiUrl);
                req.UserAgent = "EmailNotes-Updater";
                req.Accept = "application/vnd.github+json";
                req.Timeout = 8000;

                string json;
                using (var resp = (HttpWebResponse)req.GetResponse())
                using (var sr = new StreamReader(resp.GetResponseStream()))
                    json = sr.ReadToEnd();

                var data = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(json);
                string tag = Get(data, "tag_name");
                if (string.IsNullOrEmpty(tag) || !IsNewer(tag)) return null;

                string download = null;
                object assetsObj;
                if (data.TryGetValue("assets", out assetsObj) && assetsObj is ArrayList assets)
                {
                    foreach (var a in assets)
                    {
                        var ad = a as Dictionary<string, object>;
                        if (ad == null) continue;
                        string name = Get(ad, "name");
                        if (!string.IsNullOrEmpty(name) && name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                        {
                            download = Get(ad, "browser_download_url");
                            break;
                        }
                    }
                }
                return new Info { Version = tag, DownloadUrl = download, PageUrl = Get(data, "html_url") };
            }
            catch { return null; }
        }

        /// <summary>Download the installer to %TEMP% and launch it (it closes Outlook and updates).</summary>
        public static void DownloadAndRun(Info info)
        {
            try
            {
                if (!string.IsNullOrEmpty(info.DownloadUrl))
                {
                    string tmp = Path.Combine(Path.GetTempPath(), "EmailNotesSetup.exe");
                    try { ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12; } catch { }
                    using (var wc = new WebClient())
                    {
                        wc.Headers.Add("User-Agent", "EmailNotes-Updater");
                        wc.DownloadFile(info.DownloadUrl, tmp);
                    }
                    System.Diagnostics.Process.Start(tmp);
                    return;
                }
            }
            catch { }
            // Fallback: just open the release page in the browser.
            try { if (!string.IsNullOrEmpty(info.PageUrl)) System.Diagnostics.Process.Start(info.PageUrl); }
            catch { }
        }

        private static string Get(Dictionary<string, object> d, string key)
        {
            object v;
            return d != null && d.TryGetValue(key, out v) && v != null ? v.ToString() : null;
        }

        private static bool IsNewer(string tag)
        {
            Version v;
            if (!Version.TryParse(tag.TrimStart('v', 'V'), out v)) return false;
            Version cur = Assembly.GetExecutingAssembly().GetName().Version;
            int lb = v.Build < 0 ? 0 : v.Build;
            int cb = cur.Build < 0 ? 0 : cur.Build;
            if (v.Major != cur.Major) return v.Major > cur.Major;
            if (v.Minor != cur.Minor) return v.Minor > cur.Minor;
            return lb > cb;
        }
    }
}
