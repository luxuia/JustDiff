using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
        public YAMLTreeControl()
        {
            InitializeComponent();

            var scene = UnityYamlParser.ParseUnityYaml.ParseYaml("../../../../test/Cube.prefab");
            Console.WriteLine(scene.roots.Count);

            var root = new YamlGameObject() { name = "root" };
            foreach (var rootItem in scene.roots) {
                var go = new YamlGameObject() { name = rootItem.name };
                root.childs.Add(go);
                foreach (var child in root.childs)
                {
                    go.childs.Add(new YamlGameObject() { name = child.name });
                }
            }
            yamltree.Items.Add(root);
        }
    }

    public class YamlGameObject
    {
        public string name { get; set; }
        public ObservableCollection<YamlGameObject> childs {  get; set; }

        public YamlGameObject()
        {
            this.childs = new ObservableCollection<YamlGameObject>();
        }

    }
}
