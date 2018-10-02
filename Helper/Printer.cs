using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;


namespace Helper
{
    class Printer
    {
        MainWindow mainWindow;
        List<System.Windows.Controls.CheckBox> checkBoxes;
        public Printer(MainWindow mainWindow, List<System.Windows.Controls.CheckBox> checkBoxes)
        {
            this.checkBoxes = checkBoxes;
            this.mainWindow = mainWindow;
        }
        public void PrintFiles()
        {
            foreach (string path in DataContainer.filesToPrint)
            {
                if (path.ToLower().Contains(".png") || path.ToLower().Contains("jpg"))
                {
                    PrintImage(path);
                }
                else
                {
                    Print(path);
                }

                Thread.Sleep(2500);
            }
            foreach (string path in GetCheckedFiles(StaticDocs.filesToPrint))
            {

                if (path.ToLower().Contains(".png") || path.ToLower().Contains("jpg"))
                {
                    PrintImage(path);
                }
                else
                {
                    Print(path);
                }
                Thread.Sleep(2500);
            }

        }
        private void Print(string path)
        {
            Process p = new Process();
            p.StartInfo = new ProcessStartInfo()
            {
                CreateNoWindow = true,
                Verb = "print",
                FileName = path
            };
            p.Start();
            p.Close();

        }
        private void PrintImage(string path)
        {
            using (var pd = new System.Drawing.Printing.PrintDocument())
            {
                pd.PrintPage += (o, e) =>
                {
                    var img = System.Drawing.Image.FromFile(path);

                    e.Graphics.DrawImage(img, new System.Drawing.Point(50, 50));
                };
                pd.Print();
            }
        }
        private List<string> GetCheckedFiles(List<string> filesToPrint)
        {
            List<string> pathOfChecked = new List<string>();
            foreach (System.Windows.Controls.CheckBox checkBox in checkBoxes)
            {
                if (checkBox.IsChecked == true)
                {
                    foreach (string path in filesToPrint)
                    {
                        if (path.Contains(checkBox.Content.ToString()))
                        {
                            pathOfChecked.Add(path);
                        }
                    }
                }
            }
            return pathOfChecked;
        }
    }
}
