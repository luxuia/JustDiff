using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using NetDiff;

namespace ExcelMerge
{
    public class TreeNodeViewModel
    {
        public string DisplayName { get; set; }
        public string DisplayComponents { get; set; }
        public string StatusText { get; set; }
        public Brush Background { get; set; }
        public Brush Foreground { get; set; } = Brushes.Black;
        public Brush StatusForeground { get; set; } = DiffColors.TextHighlight;
        public TextDecorationCollection TextDecorations { get; set; }
        public bool IsVisible { get; set; }
        public bool IsExpanded { get; set; }
        public List<TreeNodeViewModel> Children { get; set; } = new List<TreeNodeViewModel>();
        public UnityDiffNode DiffNode { get; set; }
    }

    public class PropertyViewModel
    {
        public string Key { get; set; }
        public string DisplayValue { get; set; }
        public Brush Background { get; set; }
    }

    public partial class UnityTreeControl : UserControl
    {
        public string Tag_Side { get; set; }

        private bool _hideEqual = true;
        public bool HideEqual
        {
            get => _hideEqual;
            set
            {
                _hideEqual = value;
                RefreshTree();
            }
        }

        private List<UnityDiffNode> _diffNodes;
        private Dictionary<UnityDiffNode, TreeNodeViewModel> _nodeMap = new Dictionary<UnityDiffNode, TreeNodeViewModel>();
        private bool _suppressSelectionSync;

        public UnityTreeControl()
        {
            InitializeComponent();
        }

        public void SetDiffData(List<UnityDiffNode> diffNodes)
        {
            _diffNodes = diffNodes;
            RefreshTree();
        }

        void RefreshTree()
        {
            if (_diffNodes == null) return;
            _nodeMap.Clear();
            var viewModels = BuildViewModels(_diffNodes);
            DiffTree.ItemsSource = viewModels;
        }

        public void SelectByDiffNode(UnityDiffNode node)
        {
            if (node == null || !_nodeMap.TryGetValue(node, out var vm)) return;
            _suppressSelectionSync = true;
            SelectTreeViewItem(DiffTree, vm);
            ShowProperties(node);
            _suppressSelectionSync = false;
        }

        static bool SelectTreeViewItem(ItemsControl parent, TreeNodeViewModel target)
        {
            if (parent == null) return false;
            foreach (var item in parent.Items)
            {
                var container = parent.ItemContainerGenerator.ContainerFromItem(item) as TreeViewItem;
                if (container == null) continue;

                if (item == target)
                {
                    container.IsSelected = true;
                    container.BringIntoView();
                    return true;
                }

                container.IsExpanded = true;
                container.UpdateLayout();
                if (SelectTreeViewItem(container, target))
                    return true;
            }
            return false;
        }

        List<TreeNodeViewModel> BuildViewModels(List<UnityDiffNode> nodes)
        {
            var result = new List<TreeNodeViewModel>();
            foreach (var node in nodes)
            {
                var vm = new TreeNodeViewModel();
                vm.DiffNode = node;
                vm.DisplayName = node.DisplayName;
                vm.DisplayComponents = node.DisplayComponents;
                vm.Children = BuildViewModels(node.Children);

                bool visible = true;
                if (_hideEqual && node.Status == DiffStatus.Equal && !node.HasChanges)
                {
                    visible = false;
                }
                vm.IsVisible = visible;
                vm.IsExpanded = !_hideEqual || node.HasChanges;
                _nodeMap[node] = vm;

                bool isSrc = Tag_Side == "src";
                bool nameChanged = node.SrcNode != null && node.DstNode != null
                    && node.SrcNode.Name != node.DstNode.Name;

                switch (node.Status)
                {
                    case DiffStatus.Modified:
                        if (nameChanged)
                        {
                            var myName = isSrc ? node.SrcNode.Name : node.DstNode.Name;
                            var otherName = isSrc ? node.DstNode.Name : node.SrcNode.Name;
                            vm.DisplayName = myName;
                            vm.Background = DiffColors.ModifiedStrong;
                            vm.StatusText = $"(renamed: {otherName})";
                            vm.StatusForeground = DiffColors.TextHighlight;
                        }
                        else
                        {
                            vm.Background = DiffColors.Modified;
                            vm.StatusText = "(modified)";
                        }
                        break;
                    case DiffStatus.Deleted:
                        if (isSrc)
                        {
                            vm.Background = DiffColors.Deleted;
                            vm.StatusText = "(deleted)";
                            vm.TextDecorations = System.Windows.TextDecorations.Strikethrough;
                            vm.Foreground = DiffColors.TextMuted;
                        }
                        else
                        {
                            vm.DisplayName = "[无]";
                            vm.DisplayComponents = "";
                            vm.StatusText = "";
                            vm.Background = DiffColors.Deleted;
                            vm.Foreground = DiffColors.TextMuted;
                        }
                        break;
                    case DiffStatus.Inserted:
                        if (isSrc)
                        {
                            vm.DisplayName = "[无]";
                            vm.DisplayComponents = "";
                            vm.StatusText = "";
                            vm.Background = DiffColors.Inserted;
                            vm.Foreground = DiffColors.TextMuted;
                        }
                        else
                        {
                            vm.Background = DiffColors.InsertedStrong;
                            vm.StatusText = "(added)";
                            vm.StatusForeground = DiffColors.Inserted;
                        }
                        break;
                    default:
                        vm.Background = Brushes.Transparent;
                        vm.StatusText = "";
                        break;
                }
                result.Add(vm);
            }
            return result;
        }

        private void DiffTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is TreeNodeViewModel vm && vm.DiffNode != null)
            {
                ShowProperties(vm.DiffNode);

                if (!_suppressSelectionSync)
                {
                    var window = Window.GetWindow(this) as UnityDifferWindow;
                    window?.OnTreeNodeSelected(this, vm.DiffNode);
                }
            }
        }

        void ShowProperties(UnityDiffNode diffNode)
        {
            var items = new List<PropertyViewModel>();
            var isSrc = Tag_Side == "src";

            var changedMap = new Dictionary<string, PropertyDiff>();
            if (diffNode.ChangedProperties != null)
            {
                foreach (var prop in diffNode.ChangedProperties)
                    changedMap[prop.Key] = prop;
            }

            var node = isSrc ? diffNode.SrcNode : diffNode.DstNode;
            if (node != null)
            {
                foreach (var kv in node.Properties.OrderBy(k => k.Key))
                {
                    Brush bg = Brushes.Transparent;
                    var displayValue = kv.Value;

                    if (changedMap.TryGetValue(kv.Key, out var prop))
                    {
                        displayValue = (isSrc ? prop.SrcValue : prop.DstValue) ?? "";
                        switch (prop.Status)
                        {
                            case DiffStatus.Modified:
                                bg = DiffColors.Modified;
                                break;
                            case DiffStatus.Deleted:
                                bg = DiffColors.Deleted;
                                break;
                            case DiffStatus.Inserted:
                                bg = DiffColors.Inserted;
                                break;
                        }
                    }

                    items.Add(new PropertyViewModel
                    {
                        Key = kv.Key,
                        DisplayValue = displayValue,
                        Background = bg
                    });
                }
            }

            foreach (var prop in changedMap.Values.OrderBy(p => p.Key))
            {
                if (node != null && node.Properties.ContainsKey(prop.Key)) continue;

                var value = (isSrc ? prop.SrcValue : prop.DstValue) ?? "";
                Brush bg = prop.Status == DiffStatus.Inserted ? DiffColors.Inserted
                         : prop.Status == DiffStatus.Deleted ? DiffColors.Deleted
                         : DiffColors.Modified;

                items.Add(new PropertyViewModel
                {
                    Key = prop.Key,
                    DisplayValue = value,
                    Background = bg
                });
            }

            PropertyGrid.ItemsSource = items;
        }

        private void TreeViewItem_ExpandedCollapsed(object sender, RoutedEventArgs e)
        {
            if (_suppressExpandSync) return;
            if (e.OriginalSource is TreeViewItem tvi && tvi.DataContext is TreeNodeViewModel vm && vm.DiffNode != null)
            {
                var window = Window.GetWindow(this) as UnityDifferWindow;
                window?.OnTreeNodeExpandChanged(this, vm.DiffNode, tvi.IsExpanded);
            }
        }

        private bool _suppressExpandSync;

        public void SyncExpand(UnityDiffNode node, bool expanded)
        {
            if (!_nodeMap.TryGetValue(node, out var vm)) return;
            _suppressExpandSync = true;
            vm.IsExpanded = expanded;
            var container = FindContainer(DiffTree, vm);
            if (container != null)
                container.IsExpanded = expanded;
            _suppressExpandSync = false;
        }

        static TreeViewItem FindContainer(ItemsControl parent, TreeNodeViewModel target)
        {
            if (parent == null) return null;
            foreach (var item in parent.Items)
            {
                var container = parent.ItemContainerGenerator.ContainerFromItem(item) as TreeViewItem;
                if (container == null) continue;
                if (item == target) return container;
                container.IsExpanded = true;
                container.UpdateLayout();
                var found = FindContainer(container, target);
                if (found != null) return found;
            }
            return null;
        }

        private void DiffTree_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.C && Keyboard.Modifiers == ModifierKeys.Control)
            {
                CopySelectedNodeNames();
                e.Handled = true;
            }
        }

        void SafeSetClipboard(string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            try { Clipboard.SetDataObject(text); }
            catch { }
        }

        void CopySelectedNodeNames()
        {
            if (DiffTree.SelectedItem is TreeNodeViewModel vm && vm.DiffNode != null)
            {
                SafeSetClipboard(vm.DiffNode.DisplayName);
            }
        }

        private void CopyNodeName_Click(object sender, RoutedEventArgs e)
        {
            CopySelectedNodeNames();
        }

        private void CopyNodePath_Click(object sender, RoutedEventArgs e)
        {
            if (DiffTree.SelectedItem is TreeNodeViewModel vm && vm.DiffNode != null)
            {
                var data = vm.DiffNode.SrcNode ?? vm.DiffNode.DstNode;
                if (data != null)
                    SafeSetClipboard(data.Path);
            }
        }

        private void CopyChildStructure_Click(object sender, RoutedEventArgs e)
        {
            if (DiffTree.SelectedItem is TreeNodeViewModel vm && vm.DiffNode != null)
            {
                var sb = new StringBuilder();
                BuildStructureText(vm.DiffNode, 0, sb);
                if (sb.Length > 0)
                    SafeSetClipboard(sb.ToString());
            }
        }

        void BuildStructureText(UnityDiffNode node, int depth, StringBuilder sb)
        {
            var data = node.SrcNode ?? node.DstNode;
            if (data == null) return;
            sb.Append(new string(' ', depth * 2));
            sb.AppendLine(data.Name);
            foreach (var child in node.Children)
            {
                BuildStructureText(child, depth + 1, sb);
            }
        }

        private void DiffTree_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            var window = Window.GetWindow(this) as UnityDifferWindow;
            window?.OnTreeScrollChanged(this, e);
        }

        private void UserControl_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                var isSrc = Tag_Side == "src";
                Entrance.OnDragFile(files, isSrc);
            }
        }

        public ScrollViewer GetScrollViewer()
        {
            return Util.GetVisualChild<ScrollViewer>(DiffTree);
        }
    }
}
