using SharpSvn;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;

namespace ExcelMerge
{
    public class Entrance
    {
        const string Scheme = "xlsmerge://";
        public static List<string> _tempFiles = new List<string>();

        public static string SrcFile;
        public static string DstFile;

        public static MainWindow XLSDiffWindow = null;
        public static YAMLDifferWindow YAMLWindow = null;

        public static void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            foreach (var file in _tempFiles)
            {
                if (File.Exists(file))
                {
                    File.Delete(file);
                }
            }
        }

        public static string GetVersionFile(SvnClient client, Uri uri, long revision)
        {
            var tempDir = Path.GetTempPath();
            var filename = Path.GetFileName(uri.LocalPath);

            var file = Path.Combine(tempDir, "ExcelMerge_" + revision + "_" + filename);
            var checkoutArgs = new SvnWriteArgs() { Revision = revision };
            using (var fs = System.IO.File.Create(file))
            {
                client.Write(uri, fs, checkoutArgs);
            }

            _tempFiles.Add(file);

            return file;
        }

        public static void OnDragFile(string[] files, bool isSrc)
        {

            if (files != null && files.Any())
            {
                var file = files[0];
                if (Directory.Exists(file))
                {
                    MainWindow.instance.ShowDirectoryWindow(files, "");
                }
                else
                {
                    if (isSrc)
                    {
                        SrcFile = files[0];
                    }
                    else
                    {
                        DstFile = files[0];
                    }

                    if (files.Length > 1)
                    {
                        if (isSrc)
                        {
                            DstFile = files[1];
                        }
                        else
                        {
                            SrcFile = files[1];
                        }
                    }

                    try
                    {
                        Diff(SrcFile, DstFile);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"OnDragFile Diff 异常: {ex}");
                        MessageBox.Show($"无法对比文件: {ex.Message}", "ExcelMerge", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        public static void Diff(string file1, string file2)
        {
            SrcFile = file1;
            DstFile = file2;

            var ext = Path.GetExtension(file1);
            if (ext == ".prefab" || ext == ".scene")
            {
                if (YAMLWindow == null)
                {
                    YAMLWindow = new YAMLDifferWindow();
                    YAMLWindow.Show();
                }
                YAMLWindow.Refresh();


            } else //if (ext == ".xls" || ext == ".xlsx")
            {
                if (XLSDiffWindow == null)
                {
                    XLSDiffWindow = new MainWindow();
                    XLSDiffWindow.Show();
                }
                XLSDiffWindow.Refresh();
            }
        }

        public static void DiffUri(long revision, long revisionto, Uri uri)
        {
            using (SvnClient client = new SvnClient())
            {

                var file1 = GetVersionFile(client, uri, revision);

                var file2 = GetVersionFile(client, uri, revisionto);

                Diff(file1, file2);
            }
        }

        public static void DiffUri(long revision, Uri uri, long cmprevision, Uri cmpuri)
        {
            using (SvnClient client = new SvnClient())
            {
                var file1 = GetVersionFile(client, uri, revision);

                var file2 = GetVersionFile(client, cmpuri, cmprevision);

                Diff(file1, file2);
            }
        }

        public static void Diff(long revision, long revisionto)
        {
            Uri uri;
            using (SvnClient client = new SvnClient())
            {
                string file = SrcFile;
                SvnInfoEventArgs info;
                client.GetInfo(file, out info);
                uri = info.Uri;
            }
            DiffUri(revision, revisionto, uri);
        }

        public static void DiffList(string[] difflist)
        {
            if (difflist.Length < 2) return;

            var file = difflist[0];
            string[] vs = new string[difflist.Length - 1];
            Array.Copy(difflist, 1, vs, 0, difflist.Length - 1);

            var versions = vs.Select((r) => { return int.Parse(r); }).ToList();
            versions.Sort();

            SrcFile = file;

            var config = LoadConfig();
            var baseUrl = config?.SvnBaseUrl ?? "http://m1.svn.ejoy.com/m1/";
            if (!baseUrl.EndsWith("/")) baseUrl += "/";

            DiffUri(versions[0] - 1, versions[versions.Count - 1], new Uri(baseUrl + file));
        }

        static Config LoadConfig()
        {
            try
            {
                var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
                if (File.Exists(path))
                {
                    return Newtonsoft.Json.JsonConvert.DeserializeObject<Config>(File.ReadAllText(path));
                }
            }
            catch { }
            return null;
        }

        static bool TryParseRevisionFromUrl(string path, out string url, out long rev)
        {
            rev = 0;
            url = path ?? string.Empty;
            if (string.IsNullOrEmpty(path)) return false;
            var revisionidx = path.LastIndexOf("?revision=");
            revisionidx = revisionidx >= 0 ? revisionidx + "?revision=".Length : -1;
            if (revisionidx < 0)
            {
                revisionidx = path.LastIndexOf("?r=");
                revisionidx = revisionidx >= 0 ? revisionidx + "?r=".Length : -1;
            }
            if (revisionidx > 0)
            {
                var srev = path.Substring(revisionidx);
                var endIdx = srev.IndexOf('&');
                var revStr = endIdx >= 0 ? srev.Substring(0, endIdx) : srev;
                return long.TryParse(revStr, out rev);
            }
            return true;
        }
        public static void ProcessInput(object sender, StartupEventArgs e)
        {
            try
            {
                if (e.Args.Length > 1)
                {
                    if (e.Args[0] == "-difflist")
                    {
                        string[] input = new string[e.Args.Length - 1];
                        Array.Copy(e.Args, 1, input, 0, e.Args.Length - 1);
                        DiffList(input);
                    }
                    else
                    {
                        Diff(e.Args[0], e.Args[1]);
                    }
                }
                else if (e.Args.Length == 1)
                {
                    var url = e.Args[0];
                    if (url.StartsWith(Scheme))
                    {
                        url = url.Substring(Scheme.Length);
                        int cmpidx = url.LastIndexOf("&cmp=");
                        if (cmpidx > 0)
                        {
                        string path1 = url.Substring(0, cmpidx);
                        string fileurl = string.Empty;
                        long rev = 0;
                        string path2 = url.Substring(cmpidx + "&cmp=".Length);
                        string cmpfileurl = string.Empty;
                        long cmprev = 0;

                        if (TryParseRevisionFromUrl(path1, out fileurl, out rev) &&
                            TryParseRevisionFromUrl(path2, out cmpfileurl, out cmprev))
                        {
                            DiffUri(rev, new Uri(fileurl), cmprev, new Uri(cmpfileurl));
                        }
                        else
                        {
                            MessageBox.Show("无法解析版本号", "xlsmerge", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                        }
                    else
                    {
                        string fileurl = string.Empty;
                        long rev = 0;
                        if (TryParseRevisionFromUrl(url, out fileurl, out rev) && !string.IsNullOrEmpty(fileurl))
                        {
                            DiffUri(rev - 1, rev, new Uri(fileurl));
                        }
                        else
                        {
                            MessageBox.Show("无法解析版本号", "xlsmerge", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }
                    }
                }
                else
                {
                    Diff("", "");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ProcessInput 异常: {ex}");
                MessageBox.Show($"无法打开或对比文件: {ex.Message}", "ExcelMerge", MessageBoxButton.OK, MessageBoxImage.Error);
                Diff("", "");
            }
        }
    }
}
