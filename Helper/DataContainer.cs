using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Helper
{
    static class DataContainer
    {
        //Header
        public static string manomName { get { return "МанометрПро"; } }
        public static string hollName { get { return "Компания Холл"; } }

        public static string manomBanknumber { get { return "BY24BELB30121116940020226000"; } }
        public static string hollBanknumber { get { return "3012111950007"; } }

        public static string manomBottomLine { get { return "УНП: 790822108, ОКПО: 300671777000"; } }
        public static string hollBottomLine { get { return "УНП: 790899608"; } }
        // Executors
        public static (string name,string email) GetAdelinInfo()
        {
            return (name: "375-29-546-03-33 Булатая Аделина Владимировна", email: " 555@holl-company.com ");
        }

        //other

    }
}
