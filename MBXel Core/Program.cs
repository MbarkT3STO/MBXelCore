using MBXel_Core.Core;
using MBXel_Core.Core.Units;

using System.Linq;

using System;
using System.Collections.Generic;
using MBXel_Core.Extensions;

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

            #region Export 

            /*--------------------*/
            /* Export data        */
            /*--------------------*/
            //var exporter = new XLExporter();
            //await exporter.ExportAsync(Orders, Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX");
            //Console.WriteLine("Saved");

            /*------------------------------------------*/
            /* Export data with a custom column headers */
            /*------------------------------------------*/
            //var headers = new List<string> { "Order ID", "Customer", "Product name", "Total" };
            //var exporter = new XLExporter();
            //await exporter.ExportAsync(Orders, Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX", headers);
            //Console.WriteLine("Saved");


            /*--------------------*/
            /* Using WorkBook     */
            /*--------------------*/
            //var workbook = new Workbook(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX", 2);
            //await workbook.BuildSheetAsync(0, "MB-WAR", Orders);
            //await workbook.BuildSheetAsync(1, "Orders", Orders);
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");

            /*----------------------------------------------------*/
            /* Using WorkBook with custom column headers and name */
            /*----------------------------------------------------*/
            //var workbook = new Workbook(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX", 4);
            //var headers = new List<string> { "Order ID", "Customer", "Product name", "Total" };
            //await workbook.BuildSheetAsync(0, Orders, headers, "MB-WAR");
            //await workbook.BuildSheetAsync(2, Orders, sheetName: "Orders");
            //await workbook.BuildSheetAsync(3, Orders, columnHeaders: headers);
            //// The fourth worksheet will be created but empty because we don't send data to it
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");

            /*------------------------------------------*/
            /* Using WorkBook and set a custom password */
            /*------------------------------------------*/
            //var workbook = new Workbook(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX", 2);
            //await workbook.BuildSheetAsync(0, Orders);
            //await workbook.SetPasswordAsync("123456");
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");


            #endregion

            //---------------------------------------------------------------------------------------------------------
            // Import 
            //---------------------------------------------------------------------------------------------------------
            // NOTE : This feauture works on Windows OS only (at this moment)
            //---------------------------------------------------------------------------------------------------------

            #region Using XLImporter

            /*------------------------------*/
            /* Import a specific sheet data */
            /*------------------------------*/

            //var importer = new XLImporter();

            //var Wbook = await importer.ImportSheetAsync(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx", "MB-WAR");

            /*---------------------------------*/
            /* Use LINQ with the imported data */
            /*---------------------------------*/
            //var R = (from x in Wbook select new { ID = x["Order ID"], Client = x["Customer"], Product = x["Product name"], Total = x["Total"] }).ToList();

            //R.ForEach(x =>
            //          {
            //              Console.WriteLine($"{x.ID}, {x.Client}, {x.Product}, {x.Total}");
            //          });

            #endregion

            #region Using Workbook

            /*---------------------------------------------------*/
            /* Import a specific sheet data from loaded workbook */
            /*---------------------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var workbook = new Workbook(path);
            //workbook.LoadFromFile();

            //var firstSheetDataAsQueryable = await workbook.GetSheetAsQueryableAsync("MB-WAR");
            //var R = (from x in firstSheetDataAsQueryable select new { ID = x["Order ID"], Client = x["Customer"], Product = x["Product name"], Total = x["Total"] }).ToList();

            //R.ForEach(x =>
            //          {
            //              Console.WriteLine($"{x.ID}, {x.Client}, {x.Product}, {x.Total}");
            //          });


            /*------------------------------------------------------*/
            /* Direct querying on the worksheet with Anonymous type */
            /*------------------------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var workbook = new Workbook(path);
            //workbook.LoadFromFile();

            //var result = await workbook.SelectAsync(sheetIndex: 0, selector: x => new { ID = x["Order ID"], Client = x["Customer"], Product = x["Product name"], Total = x["Total"] });

            //result.ForEach(x =>
            //{
            //    Console.WriteLine($"{x.ID}, {x.Client}, {x.Product}, {x.Total}");
            //});

            /*-----------------------------------------------------------*/
            /* Direct querying on the worksheet with known/Specific type */
            /*-----------------------------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var workbook = new Workbook(path);
            //workbook.LoadFromFile();

            //var result = await workbook.SelectAsync(sheetName: "MB-WAR", selector: x => new Order { ID = x["Order ID"].ToInt(), Client = x["Customer"], Product = x["Product name"], Total = x["Total"].ToInt() });

            //result.ForEach(x =>
            //{
            //    Console.WriteLine($"{x.ID}, {x.Client}, {x.Product}, {x.Total}");
            //});

            #endregion

            //---------------------------------------------------------------------------------------------------------
            // Modify a workbook sheets 
            //---------------------------------------------------------------------------------------------------------
            // NOTE : This feauture works on Windows OS only (at this moment)
            //---------------------------------------------------------------------------------------------------------

            /*-----------*/
            /* Example 1 */
            /*-----------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var workbook = new Workbook(path);
            //workbook.LoadFromFile();
            ////Add new worksheet to the workbook
            //await workbook.InsertEmptySheetAsync("New Sheet");
            ////Fill the new added worksheet
            //await workbook.BuildSheetAsync(workbook.SheetsCount - 1, Orders);
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");


            /*-----------*/
            /* Example 2 */
            /*-----------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var workbook = new Workbook(path);
            //workbook.LoadFromFile();
            ////Add new worksheet to the workbook
            //await workbook.InsertEmptySheetAsync();
            ////Fill the new added worksheet
            //await workbook.BuildSheetAsync(workbook.SheetsCount - 1, Orders);
            ////Remove the first worksheet from the workbook
            //workbook.RemoveSheet(0);
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");

            //---------------------------------------------------------------------------------------------------------
            // Using chaining methods
            //---------------------------------------------------------------------------------------------------------

            /*-----------*/
            /* Example 1 */
            /*-----------*/
            //var workbook = new Workbook(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX", 4);
            //var headers = new List<string> { "Order ID", "Customer", "Product name", "Total" };
            //await workbook
            //     .BuildWorkSheet(0, Orders, headers, "MB-WAR")
            //     .BuildWorkSheet(2, Orders, sheetName: "Orders")
            //     .BuildWorkSheet(3, Orders, columnHeaders: headers)
            //     .SaveAsync();
            //Console.WriteLine("Saved");

            /*-----------*/
            /* Example 2 */
            /*-----------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var workbook = new Workbook(path);
            //await workbook.LoadFile()
            //              .InsertEmptyWorkSheet()
            //              .BuildWorkSheet( workbook.SheetsCount - 1 , Orders )
            //              .RemoveWorkSheet( 0 )
            //              .Protect( "123456" )
            //              .SaveAsync();
            //Console.WriteLine("Saved");

            Console.ReadKey();
        }
    }
}
