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

namespace Helper
{
    class Offer
    {
        string pathOfChangElement = Directory.GetCurrentDirectory() + @"\doc\changeable";
        string pathOfInitialData = @"doc\initial_data";


        MainWindow mainWindow;
        List<Product> products;
        public Offer(MainWindow mainWindow, List<Product> products)
        {
            this.mainWindow = mainWindow;
            this.products = products;


        }
        public void PrepareOffer()
        {
            CreateDocument();
        }
        public void CreateDocument()
        {
            var wordApp = new Word.Application();
            var wordDoc = wordApp.Documents.Add(pathOfChangElement + @"\offer.docx");//Открываем шаблон

            //Вставляем все данные
            SetHeader(wordDoc);
            MainWindow.ReplaceStub("{num}", mainWindow.ProcedureNumberTextBox.Text, wordDoc); //Заменяем метку на данные из формы(здесь конкретно из текстбокса с именем textBox_fio)
            MainWindow.ReplaceStub("{to_whom}", mainWindow.ClientTextBox.Text, wordDoc); //Заменяем метку на данные из формы(здесь конкретно из текстбокса с именем textBox_fio)
            Executor executor = new Executor(mainWindow);
            if (!executor.SetExecutor(wordDoc))
            {
                throw new NullReferenceException();
            }

            MainWindow.ReplaceStub("{date}", mainWindow.date.Text, wordDoc);
            EditGuarantee(wordDoc);
            SetMoney(wordDoc);
            MainWindow.ReplaceStub("{validity}", mainWindow.validity.Text, wordDoc);
            MainWindow.ReplaceStub("{terms_payment}", mainWindow.deliveryConditionsTextBox.Text, wordDoc);
            MainWindow.ReplaceStub("{delivery_time}", mainWindow.deliveryTimeTextBox.Text, wordDoc);
            MainWindow.ReplaceStub("{delivery-conditions}", mainWindow.deliveryConditionsTextBox.Text, wordDoc);
            MainWindow.ReplaceStub("{delivery_address}", mainWindow.deliveryAddressTextBox.Text, wordDoc);
            EditProducts(wordDoc);

            //Сохраняем
            string pathToSave = Directory.GetCurrentDirectory().ToString() + @"\doc\changeable\offerTmp.docx";//Добавить класс сохранения по папкам

            try
            {
                wordDoc.SaveAs2(pathToSave);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("Закройте открытый файл: offerTmp");//Исправить
                throw new NullReferenceException();
            }

            wordDoc.Close();
            if (DataContainer.filesToPrint.IndexOf(pathToSave) < 0)
            {
                DataContainer.filesToPrint.Add(pathToSave);//Изменить!! Никаких абсолютных путей
            }
        }
        public void OpenFile()
        {
            //Открываем полученный результат( тут будет рзветвление)
            Process.Start(@"doc\changeable\offerTmp.docx");
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
        public void EditGuarantee(Document wordDoc)
        {
            string guarantee = mainWindow.guaranteeTextBox.Text;
            if (guarantee.Length > 2 && guarantee != "112" && guarantee != "113" && guarantee != "114")
            { guarantee = guarantee.Substring(1); }
            else if (guarantee.Length > 1 && guarantee != "12" && guarantee != "13" && guarantee != "14")
            {
                guarantee = guarantee.Substring(1);
            }
            switch (guarantee)
            {
                case "1": { MainWindow.ReplaceStub("{guarantee}", mainWindow.guaranteeTextBox.Text + " месяц", wordDoc); break; }
                case "2":
                case "3":
                case "4": { MainWindow.ReplaceStub("{guarantee}", mainWindow.guaranteeTextBox.Text + " месяца", wordDoc); break; }
                default: { MainWindow.ReplaceStub("{guarantee}", mainWindow.guaranteeTextBox.Text + " месяцев", wordDoc); break; }
            }

        }

        private void EditProducts(Document wordDoc)
        {
            string result = "";

            foreach (Product product in products)
            {
                if (result != "")
                {
                    result += ", ";
                }
                Unit unit = new Unit(product.кол_во, product.ед_изм);
                result += product.наименование_товара + " - " + product.кол_во + " " + product.ед_изм;//unit.GetUnitInString();  Словарнная запись размерности.
            }
            MainWindow.ReplaceStub("{products}", result, wordDoc);
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
