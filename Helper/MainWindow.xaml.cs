using System;
using System.Collections.Generic;
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
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Shapes;
using System.IO;
using Microsoft.Win32;
using System.Diagnostics;
using Helper.ViewModel;
using Helper.Changeable;
using System.Printing;




namespace Helper
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        StaticDocs staticDocs;
        public MainWindow()
        {
            InitializeComponent();
            staticDocs = new StaticDocs(this);
            staticDocs.PrepareStaticFilesList();

        }
        List<Product> products = new List<Product>();
        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Directory.GetCurrentDirectory();

            if (openFileDialog.ShowDialog() == true)
            {
                DataContainer.filesToPrint.Add(openFileDialog.FileName);
            }
            

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                Specification specification = new Specification(this, products);
                specification.PrepareSpecification();
                Offer offer = new Offer(this, products);
                offer.PrepareOffer();
                specification.OpenFile();
                offer.OpenFile();
            }
            catch (NullReferenceException)
            {

            }


        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {

        }

        private void DataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {

        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            products.Clear();
            productGrid.ItemsSource = null;
            int count = 1;
            if (productCount.Value != null)
            {
                double tmp = (double)productCount.Value;
                count = (int)tmp;
            }
            productGrid.ItemsSource = CreateListForGrid(count);

        }
        private List<Product> CreateListForGrid(int count)
        {
            for (int i = 0; i < count; i++)
            {
                products.Add(new Product(i + 1, "", "", 0, "", 0.0));
            }
            return products;
        }
        public static void ReplaceStub(string stubToReplace, string text, Word.Document worldDocument)
        {
            var range = worldDocument.Content;
            range.Find.ClearFormatting();
            object wdReplaceAll = Word.WdReplace.wdReplaceAll;
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text, Replace: wdReplaceAll);
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            try
            {
                Specification specification = new Specification(this, products);
                specification.PrepareSpecification();
                Offer offer = new Offer(this, products);
                offer.PrepareOffer();

                //Редакт листа пчеати
                Printer p = new Printer(this, staticDocs.GetCheckBoxes);
                p.PrintFiles();
            }
            catch (NullReferenceException)
            {

            }//!!!!
        }
    }
}
