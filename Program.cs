using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting;
using System.Text;

namespace ExcelExport
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Clear();
            int i = 0;
            var data = new List<ListExcelData>
                       {
                           new ListExcelData {n = ++i, region = "16", person_id = "4075018", family = "Иванов", name = "Александр", fname = "Васильевич", born_on = "15.12.1968", snils = "7873591319", request_id = "5", voucher_ser = "1", voucher_number = "1", reserved_on = "22.03.2015", departure_date = "22.03.2015", price = "17200.00", cost = "5000.00"},
                           new ListExcelData {n = ++i, region = "16", person_id = "4075018", family = "Иванов", name = "Александр", fname = "Васильевич", born_on = "15.12.1968", snils = "7873591319", request_id = "5", voucher_ser = "1", voucher_number = "2", reserved_on = "25.03.2015", departure_date = "21.03.2015", price = "17200.00", cost = "7000.00"},
                           new ListExcelData {n = ++i, region = "16", person_id = "4075018", family = "Иванов", name = "Александр", fname = "Васильевич", born_on = "15.12.1968", snils = "7873591319", request_id = "5", voucher_ser = "1", voucher_number = "3", reserved_on = null, departure_date = null, price = "17200.00", cost = null},
                           new ListExcelData {n = ++i, region = "16", person_id = null, family = null, name = null, fname = null, born_on = null, snils = null, request_id = null, voucher_ser = "2", voucher_number = "1", reserved_on = null, departure_date = null, price = null, cost = null},
                           new ListExcelData {n = ++i, region = "16", person_id = null, family = null, name = null, fname = null, born_on = null, snils = null, request_id = null, voucher_ser = "2", voucher_number = "2", reserved_on = null, departure_date = null, price = null, cost = null},
                           new ListExcelData {n = ++i, region = "16", person_id = null, family = null, name = null, fname = null, born_on = null, snils = null, request_id = null, voucher_ser = "1", voucher_number = "4", reserved_on = null, departure_date = null, price = null, cost = null},
                           new ListExcelData {n = ++i, region = "16", person_id = null, family = null, name = null, fname = null, born_on = null, snils = null, request_id = null, voucher_ser = "1", voucher_number = "5", reserved_on = null, departure_date = null, price = null, cost = null},
                           new ListExcelData {n = ++i, region = "16", person_id = null, family = null, name = null, fname = null, born_on = null, snils = null, request_id = null, voucher_ser = "1", voucher_number = "6", reserved_on = null, departure_date = null, price = null, cost = null},
                           new ListExcelData {n = ++i, region = "16", person_id = null, family = null, name = null, fname = null, born_on = null, snils = null, request_id = null, voucher_ser = "1", voucher_number = "7", reserved_on = null, departure_date = null, price = null, cost = null},
                           new ListExcelData {n = ++i, region = "16", person_id = null, family = null, name = null, fname = null, born_on = null, snils = null, request_id = null, voucher_ser = "1", voucher_number = "8", reserved_on = null, departure_date = null, price = null, cost = null},
                           new ListExcelData {n = ++i, region = "16", person_id = null, family = null, name = null, fname = null, born_on = null, snils = null, request_id = null, voucher_ser = "1", voucher_number = "9", reserved_on = null, departure_date = null, price = null, cost = null},
                           new ListExcelData {n = ++i, region = "16", person_id = null, family = null, name = null, fname = null, born_on = null, snils = null, request_id = null, voucher_ser = "1", voucher_number = "10", reserved_on = null, departure_date = null, price = null, cost = null},
                           new ListExcelData {n = ++i, region = "16", person_id = null, family = null, name = null, fname = null, born_on = null, snils = null, request_id = null, voucher_ser = "3", voucher_number = "1", reserved_on = null, departure_date = null, price = null, cost = null},
                           new ListExcelData {n = ++i, region = "16", person_id = null, family = null, name = null, fname = null, born_on = null, snils = null, request_id = null, voucher_ser = "3", voucher_number = "2", reserved_on = null, departure_date = null, price = null, cost = null},
                       };

            string xlsmPath = @"D:\Морданов\C# проекты\ExcelExport\ExcelExport\bin\Debug\template.xlsm";

            try
            {
                var excel = new ExcelExporter(xlsmPath);

                excel.InsertData("voucher", data);

                var fields = new Dictionary<string, object>
                             {
                                 {"dateNow", DateTime.Now.ToString("dd.MM.yyyy")},
                                 {"contract_num", 83437},
                                 {"contract_date", new DateTime(2015, 12, 15).ToString("dd.MM.yyyy")},
                                 {"tour_from", new DateTime(2016, 4, 20).ToString("dd.MM.yyyy")},
                                 {"tour_to", new DateTime(2016, 4, 27).ToString("dd.MM.yyyy")}
                             };
                excel.InsertFields(fields);

                excel.RunMacros();
                excel.ProtectSheets("dox");
                //excel.OpenExcel();

                //excel.SaveExcel("book1.xlsm");
            }
            catch (Exception ex)
            {
                Console.WriteLine("=======================================");
                Console.WriteLine("-------Message-------");
                Console.WriteLine(ex.Message);
                Console.WriteLine("-------StackTrace-------");
                Console.WriteLine(ex.StackTrace);
            }



            Console.Write("Нажмите любую клавишу...");
            Console.ReadKey();
        }
    }

    class ListExcelData
    {
        public int n { get; set; }
        public string region { get; set; }
        public string person_id { get; set; }

        public string fiodr
        {
            get
            {
                if (person_id == null) return null;
                return String.Format("{0} {1} {2} ({3})", family, name, fname, born_on);
            }
        }
        public string family { get; set; }
        public string name { get; set; }
        public string fname { get; set; }
        public string born_on { get; set; }
        public string snils { get; set; }
        public string request_id { get; set; }
        public string voucher_ser { get; set; }
        public string voucher_number { get; set; }
        public string reserved_on { get; set; }
        public string departure_date { get; set; }
        [unprotected]
        public string reserved_on_fact { get; set; }
        [unprotected]
        public string departure_date_fact { get; set; }
        public string price { get; set; }
        public string cost { get; set; }
        [unprotected]
        public string cost_fact { get; set; }
    }

    internal class unprotectedAttribute: Attribute
    {
    }
}
