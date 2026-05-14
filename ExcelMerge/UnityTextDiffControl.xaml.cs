using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using NetDiff;

namespace ExcelMerge
{
    public class TextLineViewModel
    {
        public string LineNumber { get; set; }
        public string Text { get; set; }
        public Brush Background { get; set; }
    }

    public partial class UnityTextDiffControl : UserControl
    {
        public string Tag_Side { get; set; }

        public UnityTextDiffControl()
        {
            InitializeComponent();
        }

        public void SetDiffData(List<DiffResult<string>> diffResults)
        {
            var items = new List<TextLineViewModel>();
            bool isSrc = Tag_Side == "src";
            int lineNum = 0;

            foreach (var dr in diffResults)
            {
                if (isSrc && dr.Status == DiffStatus.Inserted)
                {
                    items.Add(new TextLineViewModel
                    {
                        LineNumber = "",
                        Text = "",
                        Background = Brushes.LightGreen
                    });
                    continue;
                }
                if (!isSrc && dr.Status == DiffStatus.Deleted)
                {
                    items.Add(new TextLineViewModel
                    {
                        LineNumber = "",
                        Text = "",
                        Background = Brushes.LightGray
                    });
                    continue;
                }

                lineNum++;
                var text = isSrc ? (dr.Obj1 ?? "") : (dr.Obj2 ?? "");

                Brush bg = Brushes.White;
                switch (dr.Status)
                {
                    case DiffStatus.Modified:
                        bg = Brushes.LightYellow;
                        break;
                    case DiffStatus.Deleted:
                        bg = Brushes.LightGray;
                        break;
                    case DiffStatus.Inserted:
                        bg = Brushes.LightGreen;
                        break;
                }

                items.Add(new TextLineViewModel
                {
                    LineNumber = lineNum.ToString(),
                    Text = text,
                    Background = bg
                });
            }

            DiffListView.ItemsSource = items;
        }

        private void DiffListView_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
        }

        public ScrollViewer GetScrollViewer()
        {
            return Util.GetVisualChild<ScrollViewer>(DiffListView);
        }
    }
}
