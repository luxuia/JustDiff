using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using NetDiff;

namespace ExcelMerge
{
    public partial class UnityDifferWindow : Window
    {
        private string _srcFile;
        private string _dstFile;
        private bool _syncing;

        public UnityDifferWindow()
        {
            InitializeComponent();

            SrcTreeControl.Tag_Side = "src";
            DstTreeControl.Tag_Side = "dst";

            PreviewKeyDown += (s, e) => { if (e.Key == Key.Escape) Close(); };
        }

        public async void Refresh(string srcFile, string dstFile)
        {
            _srcFile = srcFile;
            _dstFile = dstFile;

            SrcFilePath.Content = Path.GetFileName(srcFile);
            SrcFilePath.ToolTip = srcFile;
            DstFilePath.Content = Path.GetFileName(dstFile);
            DstFilePath.ToolTip = dstFile;

            if (string.IsNullOrEmpty(srcFile) || string.IsNullOrEmpty(dstFile))
                return;

            Title = "Unity Diff - Loading...";
            DiffProgressBar.Value = 0;
            DiffProgressBar.Visibility = Visibility.Visible;

            var progress = new Progress<double>(p =>
            {
                DiffProgressBar.Value = p * 100;
            });

            try
            {
                var result = await Task.Run(() =>
                {
                    var srcRoots = UnityDiffEngine.ParseFile(srcFile);
                    var dstRoots = UnityDiffEngine.ParseFile(dstFile);
                    var diffNodes = UnityDiffEngine.DiffRoots(srcRoots, dstRoots, progress);
                    return (srcRoots, dstRoots, diffNodes);
                });

                DiffProgressBar.Visibility = Visibility.Collapsed;

                bool hasChanges = result.diffNodes.Any(n => n.HasChanges);
                bool showAll = DiffModeCombo.SelectedIndex == 2;
                NoChangesOverlay.Visibility = (hasChanges || showAll) ? Visibility.Collapsed : Visibility.Visible;

                Title = hasChanges
                    ? $"Unity Diff - src:{result.srcRoots.Count} roots, dst:{result.dstRoots.Count} roots, diff:{result.diffNodes.Count} nodes"
                    : "Unity Diff - No Changes";

                SrcTreeControl.SetDiffData(result.diffNodes);
                DstTreeControl.SetDiffData(result.diffNodes);
            }
            catch (Exception ex)
            {
                DiffProgressBar.Visibility = Visibility.Collapsed;
                Title = $"Unity Diff - ERROR";
                MessageBox.Show($"Failed to diff:\n{ex.GetType().Name}: {ex.Message}\n\n{ex.StackTrace}",
                    "Unity Diff", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void TextDiff_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_srcFile) || string.IsNullOrEmpty(_dstFile))
                return;

            var tortoiseMerge = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                "TortoiseSVN", "bin", "TortoiseMerge.exe");

            if (!File.Exists(tortoiseMerge))
            {
                MessageBox.Show("TortoiseMerge not found at:\n" + tortoiseMerge,
                    "Unity Diff", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Process.Start(tortoiseMerge, $"/base:\"{_srcFile}\" /mine:\"{_dstFile}\"");
        }

        private void DiffModeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SrcTreeControl == null || DstTreeControl == null) return;

            int mode = DiffModeCombo.SelectedIndex;
            bool onlyNode = mode == 0;
            bool hideEqual = mode != 2;

            if (mode == 2)
                NoChangesOverlay.Visibility = Visibility.Collapsed;

            SrcTreeControl.HideEqual = hideEqual;
            DstTreeControl.HideEqual = hideEqual;

            bool needRefresh = false;
            if (UnityDiffEngine.OnlyNodeChanges != onlyNode)
            {
                UnityDiffEngine.OnlyNodeChanges = onlyNode;
                needRefresh = true;
            }

            bool ignoreDetails = onlyNode;
            if (UnityDiffEngine.IgnorePositionChanges != ignoreDetails ||
                UnityDiffEngine.IgnoreFileIdChanges != ignoreDetails)
            {
                UnityDiffEngine.IgnorePositionChanges = ignoreDetails;
                UnityDiffEngine.IgnoreFileIdChanges = ignoreDetails;
                needRefresh = true;
            }

            if (needRefresh)
                Refresh(_srcFile, _dstFile);
        }


        public void OnTreeNodeSelected(UnityTreeControl source, UnityDiffNode node)
        {
            var target = source == SrcTreeControl ? DstTreeControl : SrcTreeControl;
            target.SelectByDiffNode(node);
        }

        public void OnTreeNodeExpandChanged(UnityTreeControl source, UnityDiffNode node, bool expanded)
        {
            var target = source == SrcTreeControl ? DstTreeControl : SrcTreeControl;
            target.SyncExpand(node, expanded);
        }

        public void OnTreeScrollChanged(UnityTreeControl source, ScrollChangedEventArgs e)
        {
            if (_syncing) return;
            _syncing = true;

            var target = source == SrcTreeControl ? DstTreeControl : SrcTreeControl;
            var sv = target.GetScrollViewer();
            if (sv != null)
            {
                if (e.VerticalChange != 0)
                    sv.ScrollToVerticalOffset(e.VerticalOffset);
                if (e.HorizontalChange != 0)
                    sv.ScrollToHorizontalOffset(e.HorizontalOffset);
            }

            _syncing = false;
        }


        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Entrance.UnityDiffWindow = null;
        }
    }
}
