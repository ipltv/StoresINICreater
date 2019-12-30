using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.IO;

namespace StoresIniCreator
{
    class Store : IComparable<Store>
    {
        public int Number { get; set; }
        public string Name { get; set; }
        public string Ip { get; set; }

        public int CompareTo(Store other)
        {
            return Number.CompareTo(other.Number);
        }
    }
    class IniCreator : IDisposable
    {
        private Excel.Application application;
        private Excel.Workbook workBook;
        private Excel.Worksheet workSheet;
        private object _missingObj = System.Reflection.Missing.Value;
        private int firstDataRow;
        private List<Store> storesList;
        private bool sortFlag;
        public IniCreator(FileInfo excelFile, int firstDatarRowCount, int sortByNumber = 0)
        {
            application = new Excel.Application();
            workBook = application.Workbooks.Open(excelFile.FullName, false, true);
            workSheet = workBook.Sheets[1];
            firstDataRow = firstDatarRowCount;
            storesList = new List<Store>();
            if (sortByNumber == 1) sortFlag = true;
            else sortFlag = false;
        }
        public void CreateIniFile(string otputFileName)
        {
            FileInfo storesIniFile = new FileInfo(otputFileName);
            int i = firstDataRow;
            while (workSheet.Cells[i, 1].Value != null)
            {
                if (workSheet.Cells[i, 5].Value != null)
                {
                    storesList.Add(
                        new Store
                        {
                            Number = int.Parse(workSheet.Cells[i, 5].Value.ToString()),
                            Name = workSheet.Cells[i, 1].Value.ToString(),
                            Ip = workSheet.Cells[i, 9].Value.ToString()
                        });
                }
                i++;
            }
            if (sortFlag) storesList.Sort();
            using (StreamWriter writer = new StreamWriter(storesIniFile.FullName))
            {
                foreach (Store item in storesList)
                {
                    string ipString = $"{item.Number}\t=\t{item.Ip}";
                    string nameString = $"{item.Number}_Name={item.Name}";
                    writer.WriteLine(ipString);
                    writer.WriteLine(nameString);
                    Console.WriteLine($"Processing in Store {item.Number} {item.Name} IP {item.Ip}");
                }
            }
        }


        #region IDisposable Support
        private bool disposedValue = false; // Для определения избыточных вызовов

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: освободить управляемое состояние (управляемые объекты).
                }

                workBook.Close(false, _missingObj, _missingObj);
                application.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
                application = null;
                workBook = null;
                workSheet = null;

                disposedValue = true;
            }
        }

        // TODO: переопределить метод завершения, только если Dispose(bool disposing) выше включает код для освобождения неуправляемых ресурсов.
        ~IniCreator()
        {
            // Не изменяйте этот код. Разместите код очистки выше, в методе Dispose(bool disposing).
            Dispose(false);
        }

        // Этот код добавлен для правильной реализации шаблона высвобождаемого класса.
        public void Dispose()
        {
            // Не изменяйте этот код. Разместите код очистки выше, в методе Dispose(bool disposing).
            Dispose(true);
            // TODO: раскомментировать следующую строку, если метод завершения переопределен выше.
            GC.SuppressFinalize(this);
        }
        #endregion
    }

    class Program
    {
        static void Main(string[] args)
        {
            //Добро пожаловать!
            Console.WriteLine(Messages_ruRu.Hello);

            //Получаем файл vpn_objects.
            //Если файл существует, запускаем создание ini-файла.
            //Если файл не найден, выводим сообщение об ошибке и завершаем работу.
            FileInfo file = new FileInfo(ConfigurationManager.AppSettings.Get("vpnObjectsFilePath"));
            if (file.Exists)
            {
                IniCreator iniCreator = new IniCreator(file, int.Parse(ConfigurationManager.AppSettings.Get("firstDataRowCount")),int.Parse(ConfigurationManager.AppSettings.Get("SortedFlag")));
                iniCreator.CreateIniFile(ConfigurationManager.AppSettings.Get("OutputFileName"));
                iniCreator.Dispose();
            }
            else
            {
                Console.WriteLine(Messages_ruRu.vpnObjectsFileNotFound);
            }

            //Собираем мусор и выводим сообщение для выхода.
            System.GC.Collect();
            Console.WriteLine(Messages_ruRu.Exit);
            Console.ReadLine();         
        }
    }
}
