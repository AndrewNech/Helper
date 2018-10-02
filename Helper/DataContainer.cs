using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Helper
{
    static class DataContainer
    {
        public static List<string> filesToPrint = new List<string>();
        //Header
        public static string manomName { get { return "МанометрПро"; } }
        public static string hollName { get { return "Компания Холл"; } }

        public static string manomBanknumber { get { return "BY24BELB30121116940020226000"; } }
        public static string hollBanknumber { get { return "BY82BELB30121119500070226000"; } }

        public static string manomBanknAddress { get { return "\r\nАдрес фактический, почтовый: 212008, г.Могилёв, ул. Турова, 5"; } }
        public static string hollBanknAddress { get { return " "; } }


        public static string manomBottomLine { get { return "УНП: 790822108, ОКПО: 300671777000"; } }
        public static string hollBottomLine { get { return "УНП: 790899608"; } }
        // Executors
        public static (string name,string email) GetAdelinInfo()
        {
            return (name: "375-29-546-03-33 Булатая Аделина Владимировна", email: " 555@holl-company.com ");
        }
        public static (string name, string email) GetAngelInfo()
        {
            return (name: "8 029-541-34-59 Нестеренко Анжелика", email: " 9988@holl-company.com ");
        }
        public static (string name, string email) GetMarianInfo()
        {
            return (name: "8 044-503-50-62 Егорова Марьяна", email: " 88@holl-company.com ");
        }
        public static (string name, string email) GetMaryInfo()
        {
            return (name: "8 044-503-50-62 Гойдина Мария", email: " 18@holl-company.com ");
        }
        public static (string name, string email) GetUraInfo()
        {
            return (name: "8 044-503-50-62 Ковнацкий Юрий", email: " 81@holl-company.com ");
        }
        public static (string name, string email) GetAntonyInfo()
        {
            return (name: "8 044-503-50-64 Шуранков Антон", email: " 55555@holl-company.com ");
        }
        public static (string name, string email) GetSvetlanInfo()
        {
            return (name: "8 029-631-05-55 Гапонова Светлана", email: " 99@holl-company.com ");
        }
        public static (string name, string email) GetAlenInfo()
        {
            return (name: "8 029-631-05-55 Лисицина Алена", email: " 5588@holl-company.com ");
        }

        //other

    }
}
