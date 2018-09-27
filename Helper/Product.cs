using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Helper
{
    class Product
    {
        public int номер_лота { get; set; }
        public string наименование_товара { get; set; }
        public string ед_изм { get; set; }
        public int кол_во { get; set; }
        public string страна { get; set; }
        public double цена { get; set; }


        public Product(int number, string prooductName,string unit,int count,string country,double price)
        {
            this.номер_лота = number;
            this.наименование_товара = prooductName;
            this.ед_изм = unit;
            this.кол_во = count;
            this.страна = country;
            this.цена = price;
        }

    }
}
