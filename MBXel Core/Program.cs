
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using MBXel_Core.Core.Units;
using MBXel_Core.Enums;
using MBXel_Core.Extensions;
using Spire.Xls;
using Workbook = MBXel_Core.Core.Workbook;

namespace MBXel_Core
{
    class Program
    {
        private static readonly List<Order> Orders = new List<Order>
                                                     {
                                                         new Order( 1 ,  "Ennasiri Ali" ,    "PRD-1" , 1500 ) ,
                                                         new Order( 2 ,  "Badaoui Inas" ,    "PRD-1" , 2000 ) ,
                                                         new Order( 3 ,  "Baddouh Ali" ,     "PRD-3" , 1000 ) ,
                                                         new Order( 4 ,  "Mouslim Kawtar" ,  "PRD-2" , 3500 ) ,
                                                         new Order( 5 ,  "Essalmi Karim" ,   "PRD-1" , 2000 ) ,
                                                         new Order( 6 ,  "Nousayr Ahmed" ,   "PRD-1" , 2000 ) ,
                                                         new Order( 7 ,  "Mersaoui Fatima" , "PRD-3" , 1000 ) ,
                                                         new Order( 8 ,  "Fanar Adil" ,      "PRD-1" , 2200 ) ,
                                                         new Order( 9 ,  "Eddawdi Nawal" ,   "PRD-2" , 3200 ) ,
                                                         new Order( 10 , "Houmam Karim" ,    "PRD-1" , 2400 ) ,
                                                         new Order( 11 , "Ennasiri Ali" ,    "PRD-2" , 2000 ) ,
                                                         new Order( 12 , "Ennasiri Ali" ,    "PRD-3" , 3500 ) ,
                                                         new Order( 13 , "Essalmi Karim" ,   "PRD-2" , 1500 ) ,
                                                         new Order( 14 , "Eddawdi Nawal" ,   "PRD-1" , 2000 )
                                                     };

        public static async Task Main(string[] args)
        {

            //------------------------------------------------------------------------------------------------------------------------------
            // Examples
            //------------------------------------------------------------------------------------------------------------------------------

            //------------------------------------------------------------------------------------------------------------------------------
            // Using normal methods and classic way : This is NOT the recommended way to work with MBXel Core cause it's no longer supported
            //------------------------------------------------------------------------------------------------------------------------------

            #region Normal methods

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
            //await workbook.SetPassword("123456").SaveAsync();
            //Console.WriteLine("Saved");


            #endregion

            //---------------------------------------------------------------------------------------------------------
            // Import 
            //---------------------------------------------------------------------------------------------------------
            // NOTE : This feature works on Windows OS only (at this moment)
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


            #endregion

            //------------------------------------------------------------------------------------------------------------------------------
            // Using chaining methods : This is the recommended way to work with MBXel Core
            //------------------------------------------------------------------------------------------------------------------------------

            #region Using chaining methods

            #region Export


            /*--------------------------*/
            /* Example 0 : Normal Export*/
            /*--------------------------*/
            //var workbook = new Workbook(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\NormalExport.xlsx");
            //await workbook
            //     .BuildWorkSheet(0, Orders, sheetName: "MB-WAR")
            //     .SaveAsync();
            //Console.WriteLine("Saved"); 

            /*--------------------------*/
            /* Example 0 : Normal Export*/
            /*--------------------------*/
            //var workbook = new Workbook(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\NormalExport.xlsx"); 
            //workbook.CreateEmptyWorkSheet<Order>( "MB-WAR" ).SetData( Orders );
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");

            /*----------------------------------------*/
            /* Example 1 : Export with custom headers */
            /*----------------------------------------*/
            //var path     = Environment.GetFolderPath( Environment.SpecialFolder.Desktop ) + "\\ExportWithACustomHeaders.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version   = ExcelVersion.Version2016,
            //    Path      = path
            //};
            //var headers  = new List<string> { "Order ID", "Customer", "Product name", "Total" };
            //var workbook = new Workbook( config );
            //workbook.CreateEmptyWorkSheet<Order>( "MB-WAR" ).SetData( Orders , headers );
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");

            #endregion

            #region Edit and protect

            /*---------------------------------------*/
            /* Example 2 : Edit and protect workbook */
            /*---------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\NormalExport.xlsx";
            //var workbook = new Workbook(path);
            //await workbook.LoadFile()
            //              .InsertEmptyWorkSheet()
            //              .BuildWorkSheet(0, Orders)
            //              .RemoveWorkSheet(1)
            //              .Password("123456")
            //              .Author("MB-WAR")
            //              .SaveAsync();
            //Console.WriteLine("Saved");


            #endregion

            #region Other edits

            /*-----------*/
            /* Example 3 */
            /*-----------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version = ExcelVersion.Version2016,
            //    Path = path,
            //};
            //var workbook = new Workbook(config);
            //// Create new worksheet with a specified name and fill it with data
            //workbook.CreateEmptyWorkSheet<Order>("MB-WAR").SetData(Orders);
            //// Create new worksheet, protect it, freeze table header pane and delete the second column
            //workbook.CreateEmptyWorkSheet("Orders").SetData(Orders).Protect("123456").FreezeHeadersPane().DeleteColumn(1);
            //// Create an empty worksheet
            //workbook.CreateEmptyWorkSheet();
            //// Create an empty worksheet and fill it with data
            //workbook.CreateEmptyWorkSheet().SetData(Orders);
            //// Save the workbook
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");


            /*-----------*/
            /* Example 4 */
            /*-----------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version = ExcelVersion.Version2016,
            //    Path = path,
            //};
            //var workbook = new Workbook(config);
            //// Remove a worksheet from the workbook
            //workbook.LoadFile(path).RemoveWorkSheet("Sheet2");
            //// Edit an exist worksheet ( protect it, set its tab color, and freeze panes
            //workbook.Worksheet("MB-WAR").Protect("123456").TabColor("#2ed573").FreezeHeadersPane(0);
            //// Add new worksheet and set its data
            //workbook.CreateEmptyWorkSheet<Order>("New Orders").SetData(Orders);
            //// Change the password protection for a an already protected workbook
            //workbook.Unprotect("123456").Protect("xxxxxx", false, false);
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");

            #endregion

            #region Import data from the worksheet

            /*-----------------------------------------------------------------------------------------------------------------------*/
            /* Example 5: Get data from a worksheet as a list of specified Type */
            /*------------------------------------------------------------------*/
            /* NOTE : Use this method if the worksheet data table header titles are in the same names with TYPE (T) properties names */
            /*-----------------------------------------------------------------------------------------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version = ExcelVersion.Version2016,
            //    Path = path,
            //    // Password if the workbook to be loaded had protected by password
            //    Password = "123456"
            //};
            //var workbook = new Workbook(config);

            //// Retrieve data from the worksheet as a list of type <Order>
            //var r = await workbook.LoadFile( path ).Worksheet( "MB-WAR" ).SelectAsync<Order>();

            //foreach (var order in r)
            //{
            //    Console.WriteLine($"{order.ID}, {order.Client}, {order.Product}, {order.Total}");
            //}


            /*-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
            /* Example 5.5: Get data from a worksheet as a list of specified Type */
            /*------------------------------------------------------------------*/
            /* NOTE : Use this method if the worksheet data table header titles are in the same names with TYPE (T) properties names and contains (Worksheet) columns represent some of (T) properties */
            /*-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version = ExcelVersion.Version2016,
            //    Path = path
            //};
            //var workbook = new Workbook(config);

            //// Retrieve data from the worksheet as a list of type <Order>
            //var r = await workbook.LoadFile(path).Worksheet("MB-WAR").SelectAsync<Order>(usedPropertiesExpression: x => new { x.Client, x.Product });

            //foreach (var order in r)
            //{
            //    Console.WriteLine($"{order.Client}, {order.Product}");
            //}



            /*--------------------------------------------------------------------------------------------------------------------------------------------*/
            /* Example 6: Get data from a worksheet as list of a specified type and use a custom map between data table header titles and Type properties */
            /*------------------------------------------------------------------*/
            /* NOTE : Use this method if the worksheet data table header titles are different names with TYPE (T) properties names */
            /*--------------------------------------------------------------------------------------------------------------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version   = ExcelVersion.Version2016,
            //    Path      = path, 
            //    // Password if the workbook to be loaded had protected by password
            //    Password = "123456"
            //};
            //var workbook = new Workbook(config);
            ////// Retrieve data from the worksheet as a list of type <Order> with a custom map between data table header titles and <Order> properties
            //var r = await workbook.Worksheet("MB-WAR").SelectAsync<Order, ImplementationOfISheetColumnsMapExample>();

            //foreach (var order in r)
            //{
            //    Console.WriteLine($"{order.ID}, {order.Client}, {order.Product}, {order.Total}");
            //}


            /*-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
            /* Example 6.5: Get data from a worksheet as list of a specified type and use a custom map between data table header titles and Type properties */
            /*------------------------------------------------------------------*/
            /* NOTE : Use this method if the worksheet data table header titles are in the same names with TYPE (T) properties names and contains (Worksheet) columns represent some of (T) properties */
            /*-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXXHeaders.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version = ExcelVersion.Version2016,
            //    Path = path
            //};
            //var workbook = new Workbook(config);

            //// Retrieve data from the worksheet as a list of type <Order>
            //var r = await workbook.LoadFile(path).Worksheet("MB-WAR").SelectAsync<Order, ImplementationOfISheetColumnsMapExample>(usedPropertiesExpression: x => new { x.Client, x.Total });

            //foreach (var order in r)
            //{
            //    Console.WriteLine($"{order.Client}, {order.Total}");
            //}

            #endregion

            #region Lock cells

            /*----------------------*/
            /* Example 7 Lock cells */
            /*----------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version = ExcelVersion.Version2016,
            //    Path = path,
            //};
            //var workbook = new Workbook(config);
            //// Create a new worksheet, set its data, activate lock cells protection and allow just a range of cells to be editable when worksheet protected
            //workbook.CreateEmptyWorkSheet<Order>("MB-WAR").SetData(Orders).LockRange("A1:D15").Protect("123", SheetProtectionType.All).AllowEditRange("B1:B15");
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");

            #endregion

            #region Workbook To PDF, Image, XML and HTML

            /*--------------------------------------------------*/
            /* Example 8: Convert a workbook to PDF and save it */
            /*--------------------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version = ExcelVersion.Version2016,
            //    Path = path,
            //};
            //var workbook = new Workbook(config);
            //workbook.LoadFile();
            //// Save workbook as PDF
            //var pdfPath = Environment.GetFolderPath( Environment.SpecialFolder.Desktop ) + "\\MB.Pdf";
            //await workbook.ToPdfAsync( pdfPath );
            //Console.WriteLine( "Saved" );


            /*---------------------------------------*/
            /* Example 9: Convert a worksheet to PDF */
            /*---------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version = ExcelVersion.Version2016,
            //    Path = path,
            //};
            //var workbook = new Workbook(config);
            //workbook.LoadFile();
            //// Save a worksheet as PDF
            //var pdfPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MB-WAR.Pdf";
            //await workbook.Worksheet("MB-WAR").ToPdfAsync(pdfPath);
            //Console.WriteLine("Saved");

            /*------------------------------------------*/
            /* Example 10: Convert a worksheet to Image */
            /*------------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version = ExcelVersion.Version2016,
            //    Path = path,
            //};
            //var workbook = new Workbook(config);
            //workbook.LoadFile();
            //// Save a worksheet as Image
            //var imagePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MB-WAR.jpeg";
            //await workbook.Worksheet("MB-WAR").ToImage(imagePath);
            //Console.WriteLine("Saved");

            /*-----------------------------------------------------------*/
            /* Example 11: Convert a worksheet to a custom Image format */
            /*-----------------------------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version = ExcelVersion.Version2016,
            //    Path = path,
            //};
            //var workbook = new Workbook(config);
            //workbook.LoadFile();
            //// Save a worksheet as Image with a high quality
            //var imagePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MB-WAR.bmp";
            //await workbook.Worksheet( "MB-WAR" ).ToImage( imagePath , ImageFormat.Bmp );
            //Console.WriteLine("Saved");

            /*-----------------------------------------------*/
            /* Example 12: Convert a workbook to an XML file */
            /*-----------------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version = ExcelVersion.Version2016,
            //    Path = path,
            //};
            //var workbook = new Workbook(config);
            //workbook.LoadFile();
            //// Save a workbook as an XML
            //var xmlPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MB-WAR.Xml";
            //await workbook.ToXmlAsync( xmlPath );
            //Console.WriteLine("Saved");

            /*------------------------------------------------*/
            /* Example 13: Convert a worksheet to an HTML file */
            /*------------------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version = ExcelVersion.Version2016,
            //    Path = path,
            //};
            //var workbook = new Workbook(config);
            //workbook.LoadFile();
            //// Save a worksheet as an HTML
            //var htmlPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MB-WAR.Html";
            //await workbook.Worksheet( "MB-WAR" ).ToHtmlAsync( htmlPath );
            //Console.WriteLine("Saved");

            #endregion

            #region Grouping

            /*---------------------------*/
            /* Example 14: Group columns */
            /*---------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version = ExcelVersion.Version2016,
            //    Path = path,
            //};
            //var workbook = new Workbook(config);
            //workbook.LoadFile();
            //// Group the first two columns
            //workbook.Worksheet( "MB-WAR" ).GroupColumns( 0 , 1 );
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");

            #endregion

            #region Styling

            /*-----------------------------------------------------------*/
            /* Example 15: Custom colors for worksheet data (Used cells) */
            /*-----------------------------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var config = new WorkbookConfig()
            //             {
            //                 Extension = XLExtension.Xlsx,
            //                 Version   = ExcelVersion.Version2016,
            //                 Path      = path,
            //             };
            //var workbook = new Workbook(config);
            //workbook.LoadFile();
            //// Set custom header and body colors
            //workbook.Worksheet( "MB-WAR" ).DataHeaderColors( "#3742fa" , "#ffffff" ).DataBodyColors( "#ffa502" , "#ffffff" );
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");

            #endregion

            #region Clear a Range data

            /*-----------------------------------------------------------*/
            /* Example 16: Clear a range of cells data */
            /*-----------------------------------------------------------*/
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\XXXX.xlsx";
            //var config = new WorkbookConfig()
            //{
            //    Extension = XLExtension.Xlsx,
            //    Version = ExcelVersion.Version2016,
            //    Path = path,
            //};
            //var workbook = new Workbook(config);
            //workbook.LoadFromFile();
            //// Clear a range of cells data
            //workbook.Worksheet( "MB-WAR" ).ClearRange( "A1:A15" );
            //await workbook.SaveAsync();
            //Console.WriteLine("Saved");

            #endregion


            #region Set values for ranges

            /*---------------------------------------*/
            /* Example 17 : Set a value for a range  */
            /*---------------------------------------*/
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Classeur1";
            var workbook = new Workbook(path);
            //await workbook.LoadFile()
            //              .InsertEmptyWorkSheet()
            //              .BuildWorkSheet(0, Orders)
            //              .RemoveWorkSheet(1)
            //              .Password("123456")
            //              .Author("MB-WAR")
            //              .SaveAsync();
            
            workbook.LoadFile().Worksheet( 0 ).SetRangeValue( "A7" , "MBARK" );
            workbook.LoadFile().Worksheet( 0 ).SetRangeValue( "A10:D30" , "MB-WAR" );
            await workbook.SaveAsync();

            Console.WriteLine("Saved");


            #endregion

            #endregion

            Console.ReadKey();
        }
    }
}
