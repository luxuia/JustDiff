using NetDiff;
using System;
using System.Collections;
using System.Collections.Generic;
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
using System.Windows.Shapes;
using UnityYamlParser;
using static Org.BouncyCastle.Math.EC.ECCurve;
using static System.Reflection.Metadata.BlobBuilder;

namespace ExcelMerge
{
    /// <summary>
    /// YAMLDifferWindow.xaml 的交互逻辑
    /// </summary>
    public partial class YAMLDifferWindow : Window
    {
        public YamlDiffNode root = new YamlDiffNode() { };

        public YAMLDifferWindow()
        {
            InitializeComponent();
        }

        private void DoDiff_Click(object sender, RoutedEventArgs e)
        {
            Refresh();
        }

        private void SrcFileSheetsCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void SortKeyCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void SVNVersionBtn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void DstFileSheetsCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }


        public void Refresh()
        {
            var file1 = Entrance.SrcFile;
            var file2 = Entrance.DstFile;

            if (string.IsNullOrEmpty(file1) || string.IsNullOrEmpty(file2)) return;

            var src = UnityYamlParser.ParseUnityYaml.ParseYaml(file1);
            var dst = UnityYamlParser.ParseUnityYaml.ParseYaml(file2);
            var is_prefab = System.IO.Path.GetExtension(file1) == ".prefab";

            root = new YamlDiffNode() { };

            Action<List<GameObject>, List<GameObject>, YamlDiffNode> Compare = null;
            Compare = (List<GameObject> a, List<GameObject> b, YamlDiffNode node) =>
            {
                var option = new DiffOption<GameObject>()
                {
                    EqualityComparer = new GameObjectComparer()
                };
                a = a == null ? new List<GameObject>() : a;
                b = b == null ? new List<GameObject>() : b;

                var diff = DiffUtil.Diff(a, b, option);
                //var optimized = diff.ToList();// DiffUtil.OptimizeCaseDeletedFirst(diff);
                var optimized = DiffUtil.OptimizeCaseDeletedFirst(diff);
                optimized = DiffUtil.OptimizeCaseInsertedFirst(optimized);
                var tlist = optimized.ToList();
                optimized = DiffUtil.OptimizeShift(tlist, false);
                optimized = DiffUtil.OptimizeShift(optimized, true);
                node.diff = optimized.ToList();

                foreach (var v in node.diff)
                {
                    var dn = new YamlDiffNode();
                    node.childs.Add(dn);
                    Compare(v.Obj1?.childs, v.Obj2?.childs, dn);
                }
            };

            if (false && is_prefab)
            {
                Compare(src.roots[0].childs, dst.roots[0].childs, root);
            } else
            {
                Compare(src.roots, dst.roots, root);
            }
            // refresh ui
            SrcFilePath.Content = file1;
            DstFilePath.Content = file2;

            SrcDataGrid.RefreshData();
            DstDataGrid.RefreshData();

            Dispatcher.BeginInvoke(()=>{
                SrcDataGrid.ExpandAll();
                DstDataGrid.ExpandAll();
            });
        }

        private void Window_ContentRendered(object sender, EventArgs e)
        {
            SrcDataGrid.ExpandAll();
            DstDataGrid.ExpandAll();
        }
    }
}
