using Microsoft.Win32;
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
        public static List<string> _tempFiles = new List<string>();

        public static string SrcFile;
        public static string DstFile;

        public static MainWindow XLSDiffWindow = null;
        public static UnityDifferWindow UnityDiffWindow = null;

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

            var ext = Path.GetExtension(file1)?.ToLower();
            if (ext == ".prefab" || ext == ".scene" || ext == ".unity")
            {
                if (UnityDiffWindow == null)
                {
                    UnityDiffWindow = new UnityDifferWindow();
                    UnityDiffWindow.Show();
                }
                UnityDiffWindow.Refresh(file1, file2);
            }
            else
            {
                if (XLSDiffWindow == null)
                {
                    XLSDiffWindow = new MainWindow();
                    XLSDiffWindow.Show();
                }
                XLSDiffWindow.Refresh();
            }
        }

        static void RegisterTortoiseSvnDiffTools()
        {
            try
            {
                var exePath = Environment.ProcessPath;
                if (string.IsNullOrEmpty(exePath))
                    exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName;
                if (string.IsNullOrEmpty(exePath)) return;

                var command = $"\"{exePath}\" %base %mine";
                string[] extensions = { ".prefab", ".unity", ".scene", ".xlsx" };

                using var diffKey = Registry.CurrentUser.CreateSubKey(@"Software\TortoiseSVN\DiffTools");
                if (diffKey == null) return;
                foreach (var ext in extensions)
                {
                    var existing = diffKey.GetValue(ext) as string;
                    if (string.Equals(existing, command, StringComparison.OrdinalIgnoreCase))
                        continue;
                    diffKey.SetValue(ext, command, RegistryValueKind.String);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"RegisterTortoiseSvnDiffTools failed: {ex}");
            }
        }

        public static void ProcessInput(object sender, StartupEventArgs e)
        {
            RegisterTortoiseSvnDiffTools();
            try
            {
                if (e.Args.Length > 1)
                {
                    Diff(e.Args[0], e.Args[1]);
                }
                else if (e.Args.Length == 1)
                {
                    Diff(e.Args[0], "");
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
