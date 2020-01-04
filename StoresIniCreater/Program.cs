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
        #region Consts
        const int EXCEL_MAX_ROWS_COUNT = 1048576;
        const int EXCEL_MAX_COLUMNS_COUNT = 16384;
        #endregion
        #region Filds
        private Excel.Application application;
        private Excel.Workbook workBook;
        private Excel.Worksheet workSheet;
        private object missingObj = System.Reflection.Missing.Value;
        private int firstDataRow;
        private int storeNumberColumnCount;
        private int storeNameColumnCount;
        private int storeIPColumnCount;
        private List<Store> storesList;
        private bool sortFlag;
        #endregion
        #region Ctor
        /// <summary>
        /// Инициализирует объект для формирования ini-файла на основе данных из vpnObjects
        /// </summary>
        /// <param name="excelFile">Файл vpnObjects</param>
        /// <param name="firstDatarRowCount">номер первой строки файла, с которой необходимо начать обработку</param>
        /// <param name="storeNameColumn">Номер столбца на листе Excel, где содержится название ПБО. Отсчитывается от 1.</param>
        /// <param name="storeNumberColumn">Номер столбца на листе Excel, где содержится номер ПБО. Отсчитывается от 1.</param>
        /// <param name="workSheetCount">Номер листа в книге Excel, где содержатся данные для обработки. Отсчитывается от 1.</param>
        /// <param name="storeIPColumn">Номер столбца на листе Excel, где содержится ip-адрес шлюза ПБО. Отсчитывается от 1.</param>
        /// <param name="sortByNumber">Флаг сортировки. Если истина, то данные в ini файл записываются отсортированными по номеру ПБО.</param>
        public IniCreator(FileInfo excelFile, int firstDatarRowCount, int storeNumberColumn, int storeNameColumn, int storeIPColumn, int workSheetCount, bool sortByNumber = false)
        {
            #region Cheks
            if (excelFile == null || !excelFile.Exists) throw new ArgumentException(Messages_ruRu.iniCreatorCtorFileException,nameof(excelFile));
            if (firstDatarRowCount <= 0 || firstDatarRowCount >= EXCEL_MAX_ROWS_COUNT) throw new ArgumentException(Messages_ruRu.iniCreatorCtorNumericException, nameof(firstDatarRowCount));
            if (storeNameColumn <= 0 || storeNameColumn >= EXCEL_MAX_COLUMNS_COUNT) throw new ArgumentException(Messages_ruRu.iniCreatorCtorNumericException, nameof(storeNameColumn));
            if (storeNumberColumn <= 0 || storeNameColumn >= EXCEL_MAX_COLUMNS_COUNT) throw new ArgumentException(Messages_ruRu.iniCreatorCtorNumericException, nameof(storeNumberColumn));
            if (storeIPColumn <= 0 || storeIPColumn >= EXCEL_MAX_COLUMNS_COUNT) throw new ArgumentException(Messages_ruRu.iniCreatorCtorNumericException, nameof(storeIPColumn));
            #endregion
            #region Data initialization
            application = new Excel.Application(); //Инициализируем приложение Excel
            workBook = application.Workbooks.Open(excelFile.FullName, false, true); //Открываем книгу
            workSheet = workBook.Sheets[workSheetCount]; //Получаем ссылку на рабочий лист
            firstDataRow = firstDatarRowCount;
            storeNumberColumnCount = storeNumberColumn;
            storeNameColumnCount = storeNameColumn;
            storeIPColumnCount = storeIPColumn;
            storesList = new List<Store>();
            sortFlag = sortByNumber;
            #endregion
        }
        #endregion
        #region Methods
        public void CreateIniFile(string otputFileName)
        {
            ReadStoresData(); //Читаем данные из vpnObjects
            if (sortFlag) storesList.Sort(); //Сортируем массив с данными ПБО, если выставлен соответствующий флаг.
            WriteStoresDataToIni(otputFileName); //Записываем данные в ini
        }
        #endregion

        #region Internal Methods
        private void WriteStoresDataToIni(string otputFileName)
        {
            FileInfo storesIniFile = new FileInfo(otputFileName); //Создаем FileInfo для выходного файла
            //Записываем данные о ПБО (название/IP) в ini-файл.
            using (StreamWriter writer = new StreamWriter(storesIniFile.FullName, false, Encoding.Default))
            {
                foreach (Store item in storesList)
                {
                    string ipString = $"{item.Number}\t=\t{item.Ip}";
                    string nameString = $"{item.Number}_Name={item.Name}";
                    writer.WriteLine(ipString);
                    writer.WriteLine(nameString);
                    Console.WriteLine(Messages_ruRu.storeProcessingStatus + $"{item.Number} {item.Name} IP {item.Ip}");
                }
            }
        }
        private void ReadStoresData()
        {
            int i = firstDataRow; //Номер первой обрабатываемой строки на листе
            //Получаем номер последней заполненной строки на листе
            Excel.Range last = workSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastRowCount = last.Row;
            //Считываем данные из Excel-файла
            Console.WriteLine(Messages_ruRu.vpnObjectsFileStartRead);
            double percentOfComplete = 0; //Процент обработанных строк файла
            Timer timer = new Timer((state) =>
            {
                Console.WriteLine(Messages_ruRu.percentOfReadProcess + $"{percentOfComplete:##}");
            },null,1000,1500); //Запускаем таймер для отображения процента в консоль
            while (i <= lastRowCount)
            {
                if (workSheet.Cells[i, storeNumberColumnCount].Value != null) //Проверям, что номер ПБО заполнен
                {
                    //Переносим данные в объект Store
                    storesList.Add(
                        new Store
                        {
                            Number = int.Parse(workSheet.Cells[i, storeNumberColumnCount].Value.ToString()), //номер ПБО
                            Name = workSheet.Cells[i, storeNameColumnCount].Value.ToString(), //Название ПБО
                            Ip = workSheet.Cells[i, storeIPColumnCount].Value.ToString() // IP-адрес шлюза
                        });
                }
                percentOfComplete = (double)i / lastRowCount * 100; // пересчитываем процент
                i++; //Увеличиваем счетчик. Переходим на следующую строку.
            } //Считываем данные со всех строк файла
            timer.Dispose();
            Console.WriteLine(Messages_ruRu.vpnObjectsFileFinishRead); // Выводим сообщение о завершеннии чтения
        }
        #endregion
        #region IDisposable Support
        private bool disposedValue = false; // Для определения избыточных вызовов

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: освободить управляемое состояние (управляемые объекты).
                    storesList = null;
                }

                workBook.Close(false, missingObj, missingObj);
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
        static IniCreator iniCreator;
        static void Main(string[] args)
        {
            try
            {
                //Добро пожаловать!
                Console.WriteLine(Messages_ruRu.Hello);

                //Получаем файл vpn_objects.
                //Если файл существует, запускаем создание ini-файла.
                //Если файл не найден, выводим сообщение об ошибке и завершаем работу.
                FileInfo file = new FileInfo(ConfigurationManager.AppSettings.Get("vpnObjectsFilePath"));
                if (file.Exists)
                {
                    iniCreator = new IniCreator(file, 
                        int.Parse(ConfigurationManager.AppSettings.Get("firstDataRowCount")),
                        int.Parse(ConfigurationManager.AppSettings.Get("storeNumberColumn")),
                        int.Parse(ConfigurationManager.AppSettings.Get("storeNameColumn")),
                        int.Parse(ConfigurationManager.AppSettings.Get("storeIPColumn")),
                        int.Parse(ConfigurationManager.AppSettings.Get("workSheetCount")),
                        Convert.ToBoolean(int.Parse(ConfigurationManager.AppSettings.Get("SortedFlag"))));
                    iniCreator.CreateIniFile(ConfigurationManager.AppSettings.Get("OutputFileName"));
                }
                else
                {
                    Console.WriteLine(Messages_ruRu.vpnObjectsFileNotFound);
                }
                //Выводим сообщение для выхода.
                Console.WriteLine(Messages_ruRu.Exit);
                Console.ReadLine();
            }
            catch(UnauthorizedAccessException e)
            {
                Console.WriteLine(Messages_ruRu.unauthorizedAccessException);
                Console.WriteLine(e.ToString());
            }
            catch(ConfigurationErrorsException e)
            {
                Console.WriteLine(Messages_ruRu.configFileException);
                Console.WriteLine(e.ToString());
            }
            catch(Exception e)
            {
                Console.WriteLine(Messages_ruRu.FatalError);
                Console.WriteLine(e.ToString());
            }
            finally
            {
                iniCreator.Dispose(); //Освобождаем ресурсы.
                GC.Collect(); //Собираем мусор.
            }
       
        }
    }
}
