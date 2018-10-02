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
    class Executor
    {
        MainWindow mainWindow;
        public Executor(MainWindow mainWindow)
        {
            this.mainWindow = mainWindow;
        }
        public bool SetExecutor(Document wordDoc)
        {
            switch (mainWindow.ExecutorComboBox.Text)
            {
                case "Маша":
                    {
                        MainWindow.ReplaceStub("{executor_info}", DataContainer.GetMaryInfo().name, wordDoc);
                        MainWindow.ReplaceStub("{executor_email}", DataContainer.GetMaryInfo().email, wordDoc); return true;
                    }
                case "Аделина":
                    {
                        MainWindow.ReplaceStub("{executor_info}", DataContainer.GetAdelinInfo().name, wordDoc);
                        MainWindow.ReplaceStub("{executor_email}", DataContainer.GetAdelinInfo().email, wordDoc); return true;
                    }
                case "Светлана":
                    {
                        MainWindow.ReplaceStub("{executor_info}", DataContainer.GetSvetlanInfo().name, wordDoc);
                        MainWindow.ReplaceStub("{executor_email}", DataContainer.GetSvetlanInfo().email, wordDoc); return true;
                    }
                case "Марьяна":
                    {
                        MainWindow.ReplaceStub("{executor_info}", DataContainer.GetMarianInfo().name, wordDoc);
                        MainWindow.ReplaceStub("{executor_email}", DataContainer.GetMarianInfo().email, wordDoc); return true;
                    }
                case "Юра":
                    {
                        MainWindow.ReplaceStub("{executor_info}", DataContainer.GetUraInfo().name, wordDoc);
                        MainWindow.ReplaceStub("{executor_email}", DataContainer.GetUraInfo().email, wordDoc); return true;
                    }
                case "Алена":
                    {
                        MainWindow.ReplaceStub("{executor_info}", DataContainer.GetAlenInfo().name, wordDoc);
                        MainWindow.ReplaceStub("{executor_email}", DataContainer.GetAlenInfo().email, wordDoc); return true;
                    }
                case "Антон":
                    {
                        MainWindow.ReplaceStub("{executor_info}", DataContainer.GetAntonyInfo().name, wordDoc);
                        MainWindow.ReplaceStub("{executor_email}", DataContainer.GetAntonyInfo().email, wordDoc); return true;
                    }
                case "Анжелика":
                    {
                        MainWindow.ReplaceStub("{executor_info}", DataContainer.GetAngelInfo().name, wordDoc);
                        MainWindow.ReplaceStub("{executor_email}", DataContainer.GetAngelInfo().email, wordDoc); return true;
                    }
                default:
                    {
                        MessageBox.Show("Выберите исполнителя.");
                        return false;
                    }
            }
        }
    }
}
