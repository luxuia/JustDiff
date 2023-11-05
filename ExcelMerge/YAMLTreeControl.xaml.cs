using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using UnityYamlParser;

namespace ExcelMerge
{
 
    /// <summary>
    /// YAMLTreeControl.xaml 的交互逻辑
    /// </summary>
    public partial class YAMLTreeControl : UserControl
    {
        public bool isSrc
        {
            get { return Tag as string == "src"; }
        }

        public string selfTag
        {
            get
            {
                return Tag as string;
            }
        }

        public string otherTag
        {
            get
            {
                return isSrc ? "dst" : "src";
            }
        }
        public YAMLTreeControl()
        {
            InitializeComponent();

            yamltree.Items.Clear();
        }

        public void RefreshData()
        {
            var scene = Entrance.YAMLWindow.root;
            var root = new YamlGameObject() { name = "root" };

            var issrc = isSrc;

            Action<YamlDiffNode, YamlGameObject> AddNode = null;
            AddNode = (YamlDiffNode node, YamlGameObject tree) =>
            {
                for (var i = 0; i < node.diff.Count;  i++) { 
                    var v = node.diff[i];
                    var obj = issrc ? v.Obj1 : v.Obj2;
                    var go = new YamlGameObject() { 
                        name = obj==null?"":obj.name + "=>"+ string.Join("|", obj.comps), 
                        IsExpanded = true, 
                        brush= Util.GetColorByDiffStatus(v.Status)};
                    tree.childs.Add(go);
                    {
                        AddNode(node.childs[i], go);
                    }
                }
            };
            AddNode(scene, root);

            yamltree.ItemsSource = root.childs;

        }

        public void ExpandAll()
        {
            ExpandAllNodes(yamltree);
        }

        private void ExpandAllNodes(ItemsControl control)
        {
            if (control == null)
            {
                return;
            }

            foreach (Object item in control.Items)
            {
                var treeItem = control.ItemContainerGenerator.ContainerFromItem(item) as TreeViewItem;

                treeItem.ExpandSubtree();

                if (treeItem == null || !treeItem.HasItems)
                {
                    continue;
                }

                treeItem.IsExpanded = true;
                ExpandAllNodes(treeItem as ItemsControl);
            }
        }

        private void Grid_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                Entrance.OnDragFile(files, isSrc);
            }
        }
    }

    public class YamlGameObject
    {
        public string name { get; set; }
        public bool IsExpanded { get; set; }

        public SolidColorBrush brush { get; set; }

        public ObservableCollection<YamlGameObject> childs {  get; set; }

        public YamlGameObject()
        {
            this.childs = new ObservableCollection<YamlGameObject>();
        }

    }
}
