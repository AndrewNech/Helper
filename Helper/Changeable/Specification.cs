using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
using System.Windows;
using Helper.Abstract;

namespace Helper.Changeable
{
    class Specification
    {
        string pathOfChangElement = Directory.GetCurrentDirectory() + @"\doc\changeable";
        string pathOfInitialData = @"doc\initial_data";
        const string tmpPath = @"\doc\changeable\specificationTmp.docx";

        MainWindow mainWindow;
        List<Product> products = new List<Product>();
        public Specification(MainWindow mainWindow, List<Product> products)
        {
            this.mainWindow = mainWindow;
            this.products = products;
        }



        public void PrepareSpecification()
        {
            CreateDocument();
        }
        public void CreateDocument()
        {
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            var wordDoc = wordApp.Documents.Add(pathOfChangElement + @"\specification.docx");//Открываем шаблон

            //Вставляем все данные
            SetHeader(wordDoc);
            MainWindow.ReplaceStub("{num}", mainWindow.ProcedureNumberTextBox.Text, wordDoc); //Заменяем метку на данные из формы(здесь конкретно из текстбокса с именем textBox_fio)
            Executor executor = new Executor(mainWindow);
            if (!executor.SetExecutor(wordDoc))
            {
                throw new NullReferenceException();
            }
            MainWindow.ReplaceStub("{date}", mainWindow.date.Text, wordDoc);
            ChangeTable(wordDoc);
            SetMoney(wordDoc);
            MainWindow.ReplaceStub("{tech_description}", mainWindow.techDescriptionTextBox.Text, wordDoc);

            //Сохраняем
            try
            {
                wordDoc.SaveAs2(Directory.GetCurrentDirectory() + tmpPath);

            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("Закройте открытый файл: specificationTmp");//Исправить
                throw new NullReferenceException();
            }
            wordDoc.Close();
            if (DataContainer.filesToPrint.IndexOf(Directory.GetCurrentDirectory() + tmpPath) < 0)
            {
                DataContainer.filesToPrint.Add(Directory.GetCurrentDirectory() + tmpPath);//Изменить!! Никаких абсолютных путей!
            }

        }
        public void OpenFile()
        {
            //Открываем полученный результат( тут будет рзветвление)
            Process.Start(Directory.GetCurrentDirectory() + tmpPath);//Вынести в константу
        }
        private void ChangeTable(Document wordDoc)
        {
            WordTable table = new WordTable(products, wordDoc);
            table.EditWordTable();


        }

        public void SetMoney(Document wordDoc)
        {
            //Общая сумма
            MainWindow.ReplaceStub("{sum_with_tax}", WordTable.sumMoney.sum_with_tax.ToString("0.00"), wordDoc);
            //Общая сумма прописью
            StringBuilder result = new StringBuilder();
            decimal sum = Decimal.Parse(WordTable.sumMoney.sum_with_tax.ToString("0.00"));
            Sum.Пропись(sum, Валюта.Рубли, result);
            MainWindow.ReplaceStub("{sum_in_string}", result.ToString(), wordDoc);
            result.Clear();
            //НДС
            MainWindow.ReplaceStub("{sum_tax}", WordTable.sumMoney.sum_tax.ToString("0.00"), wordDoc);
            //Ндс прописью
            sum = Decimal.Parse(WordTable.sumMoney.sum_tax.ToString("0.00"));
            Sum.Пропись(sum, Валюта.Рубли, result);
            MainWindow.ReplaceStub("{sum_tax_in_string}", result.ToString(), wordDoc);
            result.Clear();
        }


        public void SetHeader(Document wordDoc)
        {
            if (mainWindow.manomRadioBut.IsChecked == true)
            {
                MainWindow.ReplaceStub("{company_name}", DataContainer.manomName, wordDoc);
                MainWindow.ReplaceStub("{bank_number}", DataContainer.manomBanknumber, wordDoc);
                MainWindow.ReplaceStub("{bank_info}", DataContainer.manomBottomLine, wordDoc);
                MainWindow.ReplaceStub("{bank_address}", DataContainer.manomBanknAddress, wordDoc);

            }
            else if (mainWindow.hollRadioBut.IsChecked == true)
            {
                MainWindow.ReplaceStub("{company_name}", DataContainer.hollName, wordDoc);
                MainWindow.ReplaceStub("{bank_number}", DataContainer.hollBanknumber, wordDoc);
                MainWindow.ReplaceStub("{bank_info}", DataContainer.hollBottomLine, wordDoc);
                MainWindow.ReplaceStub("{bank_address}", DataContainer.hollBanknAddress, wordDoc);

            }
            else
            {
                throw new Exception();
            }
        }
    }
}
