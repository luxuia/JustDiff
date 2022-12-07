using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using NetDiff;
using System.IO;
using System.Windows.Input;
using System.Net;

namespace ExcelMerge {
    class CellTemplateSelector : DataTemplateSelector {

        public CellTemplateSelector(string binder, int columnID, string tag) {
            Binder = binder;
            ColumnID = columnID;
            this.tag = tag;
        }

        public string Binder;
        public int ColumnID;
        public string tag;

        public override System.Windows.DataTemplate SelectTemplate(object item, System.Windows.DependencyObject container) {
            ExcelData rowdata = item as ExcelData;
            if (rowdata != null) {

                Brush bg = Brushes.White;
                    var rowdiff = rowdata.diffstatus;
                    if (rowdiff != null && rowdiff.diffcells.Count > ColumnID && ColumnID >= 0) {
                        var diffid = rowdata.column2diff[ColumnID];
                        var diffresult = rowdiff.diffcells[diffid];
                        DiffStatus status = rowdiff.diffcells[diffid].Status;

                        var diff_detail = rowdiff.diffcell_details != null ? rowdiff.diffcell_details[diffid]:null;

                        switch (status) {
                            case DiffStatus.Modified:
                                bg = Brushes.Yellow;
                                break;
                            case DiffStatus.Deleted:
                                // 列增删的时候不好处理，不显示影响的格子
                                if (rowdata.tag == "src")
                                    bg = Brushes.Gray;
                                break;
                            case DiffStatus.Inserted:
                                // 列增删的时候不好处理，不显示影响的格子
                                if (rowdata.tag == "dst")
                                    bg = Brushes.LightGreen;
                                break;
                            default:
                                break;
                        }
                        if (diff_detail != null && diff_detail.Count > 1) {
                            FrameworkElementFactory stackPanel = new FrameworkElementFactory(typeof(StackPanel));
                            stackPanel.SetValue(StackPanel.OrientationProperty, Orientation.Horizontal);
                            stackPanel.SetValue(StackPanel.BackgroundProperty, bg);
                            for (int i = 0; i < diff_detail.Count; ++i) {
                                if (diff_detail[i] != null) {
                                    
                                    var diff_cell = diff_detail[i];
                                    if (diff_cell.Status == DiffStatus.Deleted && tag == "dst")
                                        continue;
                                    if (diff_cell.Status == DiffStatus.Inserted && tag == "src")
                                        continue;
                                    FrameworkElementFactory textBlock = new FrameworkElementFactory(typeof(TextBlock));
                                    var text = tag == "dst" ? diff_cell.Obj2.ToString() : diff_cell.Obj1.ToString();
                                    textBlock.SetValue(TextBlock.TextProperty, text);
                                   
                                    stackPanel.AppendChild(textBlock);
                                    if (diff_cell.Status == DiffStatus.Deleted) {
                                        textBlock.SetValue(TextBlock.TextDecorationsProperty, TextDecorations.Strikethrough);
                                    } else if (diff_cell.Status == DiffStatus.Inserted) {
                                        textBlock.SetValue(TextBlock.TextDecorationsProperty, TextDecorations.Underline);
                                    }
                                    if (diff_cell.Status != DiffStatus.Equal) {
                                        textBlock.SetValue(TextBlock.ForegroundProperty, System.Windows.Media.Brushes.Red);
                                    }
                                }
                            }

                            return new DataTemplate() { VisualTree = stackPanel };
                        } else {
                            FrameworkElementFactory textBlock = new FrameworkElementFactory(typeof(TextBlock));
                            textBlock.SetValue(TextBlock.BackgroundProperty, bg);
                            textBlock.SetValue(TextBlock.TextProperty, rowdata.data[Binder].value);
                            return new DataTemplate() { VisualTree = textBlock };
                        }
                    }
            }

            return new DataTemplate() { VisualTree = new FrameworkElementFactory(typeof(TextBlock)) };
        }
    }
}
