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
    //string path = "";
    public class Offer : IDocument
    {
        string pathOfChangElement = Directory.GetCurrentDirectory() + @"\doc\changeable";
        string pathOfInitialData = @"doc\initial_data";


        MainWindow mainWindow;
        public Offer(MainWindow mainWindow)
        {
            this.mainWindow = mainWindow;

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
            ReplaceStub("{num}", mainWindow.ProcedureNumberTextBox.Text, wordDoc); //Заменяем метку на данные из формы(здесь конкретно из текстбокса с именем textBox_fio)
            ReplaceStub("{to_whom}", mainWindow.ClientTextBox.Text, wordDoc); //Заменяем метку на данные из формы(здесь конкретно из текстбокса с именем textBox_fio)
            SetExecutor(wordDoc);
            ReplaceStub("{date}", mainWindow.date.Text, wordDoc);
            EditGuarantee(wordDoc);
            SetMoney(wordDoc);

            //Сохраняем
            wordDoc.SaveAs2(@"D:\Job\Манометр\Helper\Helper\bin\Debug\doc\changeable\offerTmp.docx");
            wordDoc.Close();

            //Открываем полученный результат( тут будет рзветвление)
            Process.Start(@"D:\Job\Манометр\Helper\Helper\bin\Debug\doc\changeable\offerTmp.docx");
        }

        public void SetMoney(Document wordDoc)
        {
            //Общая сумма
            ReplaceStub("{sum_with_tax}", WordTable.sumMoney.sum_with_tax.ToString("0.00"), wordDoc);
            //Общая сумма прописью
            StringBuilder result = new StringBuilder();
            decimal sum = Decimal.Parse(WordTable.sumMoney.sum_with_tax.ToString("0.00"));
            Sum.Пропись(sum, Валюта.Рубли, result);
            ReplaceStub("{sum_in_string}", result.ToString(), wordDoc);
            result.Clear();
            //НДС
            ReplaceStub("{sum_tax}", WordTable.sumMoney.sum_tax.ToString("0.00"), wordDoc);
            //Ндс прописью
            sum = Decimal.Parse(WordTable.sumMoney.sum_tax.ToString("0.00"));
            Sum.Пропись(sum, Валюта.Рубли, result);
            ReplaceStub("{sum_tax_in_string}", result.ToString(), wordDoc);
            result.Clear();
        }
        public void EditGuarantee(Document wordDoc)
        {
            string guarantee = mainWindow.guaranteeTextBox.Text;
            if (guarantee.Length > 2 && guarantee!="112" && guarantee != "113" && guarantee != "114")
            { guarantee = guarantee.Substring(1); }
            else if (guarantee.Length > 1 && guarantee != "12" && guarantee != "13" && guarantee != "14")
            {
                guarantee = guarantee.Substring(1);
            }
            switch (guarantee)
            {
                case "1": { ReplaceStub("{guarantee}", mainWindow.guaranteeTextBox.Text+" месяц", wordDoc);break; }
                case "2": 
                case "3": 
                case "4": { ReplaceStub("{guarantee}", mainWindow.guaranteeTextBox.Text + " месяца", wordDoc); break; }
                default: { ReplaceStub("{guarantee}", mainWindow.guaranteeTextBox.Text + " месяцев", wordDoc); break; }
            }

            
        }
        public void SetExecutor(Document wordDoc)
        {
            switch (mainWindow.ExecutorComboBox.Text)
            {
                case "Маша":
                    {
                        ReplaceStub("{executor_info}", "нет данных", wordDoc);
                        ReplaceStub("{executor_email}", "нет данных", wordDoc); break;
                    }
                case "Аделина":
                    {
                        ReplaceStub("{executor_info}", DataContainer.GetAdelinInfo().name, wordDoc);
                        ReplaceStub("{executor_email}", DataContainer.GetAdelinInfo().email, wordDoc); break;
                    }
                case "Светлана":
                    {
                        ReplaceStub("{executor_info}", "нет данных", wordDoc);
                        ReplaceStub("{executor_email}", "нет данных", wordDoc); break;
                    }
                default:
                    {
                        MessageBox.Show("Выберите исполнителя.");
                        throw new Exception();
                    }
            }
        }
        public void SetHeader(Document wordDoc)
        {
            if (mainWindow.manomRadioBut.IsChecked == true)
            {
                ReplaceStub("{company_name}", DataContainer.manomName, wordDoc);
                ReplaceStub("{bank_number}", DataContainer.manomBanknumber, wordDoc);
                ReplaceStub("{bank_info}", DataContainer.manomBottomLine, wordDoc);
            }
            else if (mainWindow.hollRadioBut.IsChecked == true)
            {
                ReplaceStub("{company_name}", DataContainer.hollName, wordDoc);
                ReplaceStub("{bank_number}", DataContainer.hollBanknumber, wordDoc);
                ReplaceStub("{bank_info}", DataContainer.hollBottomLine, wordDoc);
            }
            else
            {
                throw new Exception();
            }
        }
        private void ReplaceStub(string stubToReplace, string text, Word.Document worldDocument)
        {
            var range = worldDocument.Content;
            range.Find.ClearFormatting();
            object wdReplaceAll = Word.WdReplace.wdReplaceAll;
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text, Replace: wdReplaceAll);
        }
    }
}
