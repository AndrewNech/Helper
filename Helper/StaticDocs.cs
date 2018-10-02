using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Helper
{
    class StaticDocs
    {
        MainWindow mainWindow;
        List<CheckBox> checkBoxes = new List<CheckBox>();
        public static List<string> filesToPrint = new List<string>();
        string path = @"\doc\static";
        public StaticDocs(MainWindow mainWindow)
        {
            this.mainWindow = mainWindow;
        }
        public List<CheckBox> GetCheckBoxes { get { return checkBoxes; } }
        public void PrepareStaticFilesList()
        {
            foreach(string file in Directory.GetFiles(Directory.GetCurrentDirectory() + path))
            {
                checkBoxes.Add(new CheckBox() { Content = file.Substring(file.LastIndexOf('\\')+1), IsChecked = true });
                mainWindow.stackFiles.Children.Add(checkBoxes.Last());
                string partialPAth = file.Substring(Directory.GetCurrentDirectory().Length+1);
                filesToPrint.Add(partialPAth);
            }
            
        }
    }
}
