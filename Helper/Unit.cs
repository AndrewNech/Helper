using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Helper
{
    class Unit
    {
        string count;
        string unit;
        public Unit(int count, string unit)
        {
            this.count = count.ToString();
            this.unit = unit;
        }
        enum doodah { штука, штуки, штук };
        enum liter { литр, литра, литров };
        enum ton { тонна, тонны, тонн };
        enum killo { килограмм, килограмма };
        enum pac { упаковка, упаковки, упаковок };





        public string GetUnitInString()
        {
            if (count.Length > 2 && count != "112" && count != "113" && count != "114")
            { count = count.Substring(1); }
            else if (count.Length > 1 && count != "12" && count != "13" && count != "14")
            {
                count = count.Substring(1);
            }
            switch (count)
            {
                case "1": { return GEtEndingFor1(); }
                case "2":
                case "3":
                case "4": { return GEtEndingFor2_4(); }
                default: { return GEtEndingForOther(); }
            }
        }
        private string GEtEndingFor1()
        {
            if (unit.ToLower().Contains("шт"))
            {
                return doodah.штука.ToString();
            }
            else if (unit.ToLower().Contains("лит"))
            {
                return liter.литр.ToString();

            }
            else if (unit.ToLower().Contains("тон") || unit.ToLower()=="")
            {
                return ton.тонна.ToString();

            }
            else if (unit.ToLower().Contains("кг"))
            {
                return killo.килограмм.ToString();
            }
            else
            {
                return doodah.штука.ToString();
            }
        }
        private string GEtEndingFor2_4()
        {
            if (unit.ToLower().Contains("шт"))
            {
                return doodah.штуки.ToString();
            }
            else if (unit.ToLower().Contains("лит"))
            {
                return liter.литра.ToString();

            }
            else if (unit.ToLower().Contains("тон"))
            {
                return ton.тонны.ToString();

            }
            else if (unit.ToLower().Contains("кг"))
            {
                return killo.килограмма.ToString();

            }
            else
            {
                return doodah.штуки.ToString();
            }
        }
        private string GEtEndingForOther()
        {
            if (unit.ToLower().Contains("шт"))
            {
                return doodah.штук.ToString();
            }
            else if (unit.ToLower().Contains("лит"))
            {
                return liter.литров.ToString();

            }
            else if (unit.ToLower().Contains("тон"))
            {
                return ton.тонна.ToString();

            }
            else if (unit.ToLower().Contains("кг"))
            {
                return killo.килограмм.ToString();

            }
            else
            {
                return doodah.штуки.ToString();
            }
        }
    }
}
