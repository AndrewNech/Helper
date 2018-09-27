using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.IO;


namespace Helper
{
    class WordTable
    {
        List<Product> products;
        List<(double sum_without_tax, double sum_tax, double sum_with_tax)> money = new List<(double sum_without_tax, double sum_tax, double sum_with_tax)>();
        Document wordDoc;
       public static (double sum_without_tax, double sum_tax, double sum_with_tax) sumMoney;
        public WordTable(List<Product> products, Document wordDoc)
        {
            this.products = products;
            this.wordDoc = wordDoc;
        }
        public (double sum_without_tax, double sum_tax, double sum_with_tax) getMoney { get {return GetSumMoney(); } }
        public void EditWordTable()
        {
            Microsoft.Office.Interop.Word.Table table = wordDoc.Tables[1];
            for (int i = 0; i < products.Count; i++)
            {
                table.Rows.Add();
                Word.Range cell = table.Cell(i + 3, 1).Range;
                cell.Text = products[i].номер_лота.ToString();

                cell = table.Cell(i + 3, 2).Range;
                cell.Text = products[i].наименование_товара.ToString();

                cell = table.Cell(i + 3, 3).Range;
                cell.Text = products[i].ед_изм.ToString();

                cell = table.Cell(i + 3, 4).Range;
                cell.Text = products[i].кол_во.ToString();

                cell = table.Cell(i + 3, 5).Range;
                cell.Text = products[i].страна.ToString();

                cell = table.Cell(i + 3, 6).Range;
                cell.Text = (products[i].цена).ToString("0.00");

                cell = table.Cell(i + 3, 7).Range;
                double sum = products[i].цена * products[i].кол_во;
                cell.Text = (sum).ToString("0.00");

                cell = table.Cell(i + 3, 8).Range;
                double sum_tax = sum * 0.2;
                cell.Text = (sum_tax).ToString("0.00");

                cell = table.Cell(i + 3, 9).Range;
                cell.Text = (sum + sum_tax).ToString("0.00");
                
                money.Add((sum, sum_tax, sum + sum_tax));
            }
            table.Rows.Add();

            table.Rows[products.Count+3].Cells[1].Merge(table.Rows[products.Count + 3].Cells[5]);//Объединение для "Итого"
            Word.Range lowerCell = table.Cell(products.Count + 3, 1).Range;//ВВод итого
            lowerCell.Text = "ИТОГО:";

            var moneyTuple = GetSumMoney();

            lowerCell = table.Cell(products.Count + 3, 3).Range;
            lowerCell.Text = moneyTuple.sum_without_tax.ToString("0.00");

            lowerCell = table.Cell(products.Count + 3, 4).Range;
            lowerCell.Text = moneyTuple.sum_tax.ToString("0.00");

            lowerCell = table.Cell(products.Count + 3, 5).Range;
            lowerCell.Text = moneyTuple.sum_with_tax.ToString("0.00");
        }
        private (double sum_without_tax, double sum_tax, double sum_with_tax) GetSumMoney()
        {
            double sum = 0;
            double sumTax = 0;
            double sumWith_tax = 0;

            for(int i = 0; i < money.Count; i++)
            {
                sum += money[i].sum_without_tax;
                sumTax += money[i].sum_tax;
                sumWith_tax += money[i].sum_with_tax;
            }
            sumMoney = (sum, sumTax, sumWith_tax);
            return sumMoney;
        }

    }
}
