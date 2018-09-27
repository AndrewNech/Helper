using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Helper.Abstract
{
    interface IDocument
    {
        void CreateDocument();
        void SetHeader(Document wordDoc);
        void SetMoney(Document wordDoc);
        void SetExecutor(Document wordDoc);
    }
}
