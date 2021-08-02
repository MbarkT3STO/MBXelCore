using MBXel_Core.Core;
using MBXel_Core.Core.Units;

using System.Linq;

using System;
using System.Collections.Generic;

namespace MBXel_Core
{
    class Program
    {
        private static readonly List<Order> Orders = new List<Order>
                                                     {
                                                         new Order(1, "Ennasiri Ali", "PRD-1", 1500),
                                                         new Order(2, "Badaoui Inas", "PRD-1", 2000),
                                                         new Order(3, "Baddouh Ali", "PRD-3", 1000),
                                                         new Order(4, "Mouslim Kawtar", "PRD-2", 3500),
                                                         new Order(5, "Essalmi Karim", "PRD-1", 2000),
                                                         new Order(6, "Nousayr Ahmed", "PRD-1", 2000),
                                                         new Order(7, "Mersaoui Fatima", "PRD-3", 1000),
                                                         new Order(8, "Fanar Adil", "PRD-1", 2200),
                                                         new Order(9, "Eddawdi Nawal", "PRD-2", 3200),
                                                         new Order(10, "Houmam Karim", "PRD-1", 2400),
                                                         new Order(11, "Ennasiri Ali", "PRD-2", 2000),
                                                         new Order(12, "Ennasiri Ali", "PRD-3", 3500),
                                                         new Order(13, "Essalmi Karim", "PRD-2", 1500),
                                                         new Order(14, "Eddawdi Nawal", "PRD-1", 2000)
                                                     };

        static async System.Threading.Tasks.Task Main(string[] args)
        {

            //---------------------------------------------------------------------------------------------------------
            // Examples
            //---------------------------------------------------------------------------------------------------------

            //---------------------------------------------------------------------------------------------------------
            // Export
            //---------------------------------------------------------------------------------------------------------

            /*--------------------*/
            /* Export data*/
            /*--------------------*/
            //var exporter = new XLExporter();
            //await exporter.ExportAsync(Orders, Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX");
            //Console.WriteLine("Saved");

            /*-----------------------------------------*/
            /* Export data with a custom column headers*/
            /*-----------------------------------------*/
            //var headers = new List<string> { "Order ID", "Customer", "Product name", "Price" };
            //var exporter = new XLExporter();
            //await exporter.ExportAsync(Orders, Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX", headers);
            //Console.WriteLine("Saved");


            /*--------------------*/
            /* Using WorkBook*/
            /*--------------------*/
            //var workbook = new Workbook(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX",2);
            //await workbook.CreateSheetAsync(0, "MB-WAR", Orders);
            //await workbook.CreateSheetAsync(1, "Orders", Orders);
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");

            /*---------------------------------------------------*/
            /* Using WorkBook with custom column headers and name*/
            /*---------------------------------------------------*/
            //var workbook = new Workbook(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX", 4);
            //var headers = new List<string> { "Order ID", "Customer", "Product name", "Price" };
            //await workbook.CreateSheetAsync(0, Orders, headers, "MB-WAR");
            //await workbook.CreateSheetAsync(2, Orders, sheetName:"Orders");
            //await workbook.CreateSheetAsync(3, Orders, columnHeaders: headers);
            //// The fourth worksheet will be created but empty because we don't send data to it
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");

            /*-----------------------------------------*/
            /* Using WorkBook and set a custom password*/
            /*-----------------------------------------*/
            //var workbook = new Workbook(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX", 2);
            //await workbook.CreateSheetAsync(0, Orders);
            //await workbook.SetPasswordAsync("123456");
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");


            //---------------------------------------------------------------------------------------------------------
            // Import 
            //---------------------------------------------------------------------------------------------------------
            // NOTE : This feauture works on Windows OS only (this moment)
            //---------------------------------------------------------------------------------------------------------

            /*-----------------------------*/
            /* Import a specific sheet data*/
            /*-----------------------------*/

            //var importer = new XLImporter();

            //var Wbook = await importer.ImportSheetAsync(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx", "MB-WAR");

            /*--------------------------------*/
            /* Use LINQ with the imported data*/
            /*--------------------------------*/
            //var R = (from x in Wbook select new { ID = x["Order ID"], Client = x["Customer"], Product = x["Product name"], Price = x["Price"] }).ToList();

            //R.ForEach(x =>
            //          {
            //              Console.WriteLine($"{x.ID}, {x.Client}, {x.Product}, {x.Price}");
            //          });

            Console.ReadKey();
        }
    }
}
