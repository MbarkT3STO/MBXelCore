using LinqToExcel;

using MBXel_Core.Core.Abstraction;

using Microsoft.AspNetCore.Http;

using OfficeOpenXml;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace MBXel_Core.Core
{
    /// <summary>
    /// Import data on the fly from an Excel file
    /// </summary>
    public class XLImporter
    {

        #region Constructors

        public XLImporter()
        {
            // This license part for EPPLUS
            // If you are a commercial business and have
            // purchased commercial licenses use the static property
            // LicenseContext of the ExcelPackage class:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        #endregion

        #region Private methods

        private Dictionary<string , int> GetSheetHeadersIndexes<T>(ExcelWorksheet worksheet, int headersRowIndex)
        {
            var result = new Dictionary<string , int>();

            var propsCountOfT = typeof(T).GetProperties().Length;

            var headers = new List<string>();

            for (int i = 1; i < propsCountOfT; i++)
            {
                if (worksheet.Cells[headersRowIndex, i].Value != null)
                {
                    var currentTitle = worksheet.Cells[headersRowIndex, i].Value.ToString();

                    if (currentTitle.Trim().Length != 0)
                    {
                        result.Add( currentTitle , i );
                    }
                }
            }

            return result;
        }
        private List<string> GetSheetColumnHeaders<T>(ExcelWorksheet worksheet, int headersRowIndex)
        {
            var propsCountOfT = typeof(T).GetProperties().Length;

            var headers = new List<string>();

            for (int i = 1; i < propsCountOfT; i++)
            {
                if (worksheet.Cells[headersRowIndex, i].Value != null)
                {
                    var currentTitle = worksheet.Cells[headersRowIndex, i].Value.ToString();
                    if (currentTitle.Trim().Length != 0)
                    {
                        headers.Add(currentTitle);
                    }
                }
            }

            return headers;

        }

        private List<T> GetDataFromSheet<T>(ExcelWorksheet worksheet, List<string> columnHeaders , int dataStartFromRow) where T : new()
        {
            var data               = new List<T>();
            var propertiesOfT      = typeof(T).GetProperties();
            var worksheetRowsCount = worksheet.Dimension.Rows;

            for (int row = dataStartFromRow; row <= worksheetRowsCount; row++)
            {
                var obj = new T();

                foreach (var prop in propertiesOfT)
                {
                    if (columnHeaders.Contains(prop.Name))
                    {
                        var currentCellValue = worksheet.Cells[row , columnHeaders.IndexOf( prop.Name ) + 1]?.ToString();
                        obj.GetType().GetProperty( prop.Name )?.SetValue( obj , Convert.ChangeType( currentCellValue , prop.PropertyType ) );
                    }
                }

                data.Add(obj);
            }

            return data;
        }
        private List<T> GetDataFromSheet<T, TSheetColumnsMap>(ExcelWorksheet worksheet , int dataStartFromRow) where T : class , new() where  TSheetColumnsMap : ISheetColumnsMap<T> , new()
        {
            var data               = new List<T>();
            var propertiesOfT      = typeof(T).GetProperties();
            var worksheetRowsCount = worksheet.Dimension.Rows;
            var headersIndexes     = GetSheetHeadersIndexes<T>( worksheet : worksheet , headersRowIndex : 1 );
            var columnsMap         = new TSheetColumnsMap().CreateMap();

            for (int row = dataStartFromRow; row <= worksheetRowsCount; row++)
            {
                var obj = new T();

                foreach (var prop in propertiesOfT)
                {
                    if (columnsMap.ContainsKey(prop.Name))
                    {
                        // Get current property related Header/Column in the worksheet
                        var propHeader = columnsMap[prop.Name];

                        // Get current property's type
                        var typeOfProp = prop.GetType();

                        // Get current property related header Index in the worksheet
                        var propHeaderIndex = headersIndexes[propHeader];

                        // Set T object current property value
                        var currentCellValue = worksheet.Cells[row, propHeaderIndex].Value.ToString(); 
                        obj.GetType().GetProperty(prop.Name)?.SetValue(obj, Convert.ChangeType(currentCellValue, prop.PropertyType));
                    }
                }
            }

            return data;
        }

        private IQueryable<Row> _Import(string filePath, string sheetName)
        {
            //Load the workbook
            using (var workbook = new ExcelQueryFactory(filePath) { ReadOnly = true})
            {
                //Collect data from the worksheet
                var result = workbook.Worksheet(sheetName);

                return result;
            }
        }
        private IQueryable<Row> _Import(string filePath, int sheetIndex)
        {
            //Load the workbook
            using (var workbook = new ExcelQueryFactory(filePath) { ReadOnly = true })
            {
                //Collect data from the worksheet
                var result = workbook.Worksheet(sheetIndex);

                return result;
            }
        }

        private Task<List<T>> _ImportFromIFormFileAsync<T>( IFormFile file, string sheetName ) where T : new()
        {
            return Task.Factory.StartNew( () =>
                                          {
                                              using ( var stream = new MemoryStream() )
                                              {
                                                  file.CopyTo( stream );

                                                  using ( var package = new ExcelPackage( stream ) )
                                                  {
                                                      var worksheet = package.Workbook.Worksheets[sheetName];
                                                      var columnHeaders = GetSheetColumnHeaders<T>( worksheet: worksheet, headersRowIndex: 1 );
                                                      var data = GetDataFromSheet<T>( worksheet , columnHeaders , 2 );

                                                      return data;
                                                  }
                                              }
                                          } );
        }
        private Task<List<T>> _ImportFromIFormFileAsync<T>(IFormFile file, int sheetIndex) where T : new()
        {
            return Task.Factory.StartNew(() =>
                                         {
                                             using (var stream = new MemoryStream())
                                             {
                                                 file.CopyTo(stream);

                                                 using (var package = new ExcelPackage(stream))
                                                 {
                                                     var worksheet = package.Workbook.Worksheets[sheetIndex];
                                                     var columnHeaders = GetSheetColumnHeaders<T>(worksheet: worksheet, headersRowIndex: 1);
                                                     var data = GetDataFromSheet<T>(worksheet, columnHeaders, 2);

                                                     return data;
                                                 }
                                             }
                                         });
        }
        private Task<List<T>> _ImportFromIFormFileAsync<T, TSheetColumnsMap>( IFormFile file, string sheetName ) where T : class, new() where TSheetColumnsMap : ISheetColumnsMap<T> , new()
        {
            return Task.Factory.StartNew( () =>
                                          {
                                              using ( var stream = new MemoryStream() )
                                              {
                                                  file.CopyTo( stream );

                                                  using ( var package = new ExcelPackage( stream ) )
                                                  {
                                                      var worksheet     = package.Workbook.Worksheets[sheetName];
                                                      var columnHeaders = GetSheetColumnHeaders<T>( worksheet: worksheet, headersRowIndex: 1 );
                                                      var data          = GetDataFromSheet<T, TSheetColumnsMap>( worksheet , 2 );

                                                      return data;
                                                  }
                                              }
                                          } );
        }
        private Task<List<T>> _ImportFromIFormFileAsync<T, TSheetColumnsMap>( IFormFile file, int sheetIndex ) where T : class, new() where TSheetColumnsMap : ISheetColumnsMap<T> , new()
        {
            return Task.Factory.StartNew( () =>
                                          {
                                              using ( var stream = new MemoryStream() )
                                              {
                                                  file.CopyTo( stream );

                                                  using ( var package = new ExcelPackage( stream ) )
                                                  {
                                                      var worksheet     = package.Workbook.Worksheets[sheetIndex];
                                                      var columnHeaders = GetSheetColumnHeaders<T>( worksheet: worksheet, headersRowIndex: 1 );
                                                      var data          = GetDataFromSheet<T, TSheetColumnsMap>( worksheet , 2 );

                                                      return data;
                                                  }
                                              }
                                          } );
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Load an excel file (Sheet) data
        /// </summary>
        /// <param name="filePath">The Excel file path</param>
        /// <param name="sheetName">The Worksheet name to select data from</param>
        /// <returns><see cref="IQueryable{Row}"/></returns>
        public IQueryable<Row> ImportSheet(string filePath, string sheetName)
        {
            return _Import(filePath, sheetName);
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="ImportSheet(string, string)"/>
        /// </summary>
        /// <inheritdoc cref="ImportSheet(string, string)"/>
        /// <returns><see cref="Task{IQueryable{Row}}"/></returns>
        public Task<IQueryable<Row>> ImportSheetAsync(string filePath, string sheetName)
        {
            return Task.Factory.StartNew(() => _Import(filePath, sheetName));
        }


        /// <summary>
        /// Load an excel file (Sheet) data
        /// </summary>
        /// <param name="filePath">The Excel file path</param>
        /// <param name="sheetIndex">The Worksheet index</param>
        /// <returns><see cref="IQueryable{Row}"/></returns>
        public IQueryable<Row> ImportSheet(string filePath, int sheetIndex)
        {
            return _Import(filePath, sheetIndex);
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="ImportSheet(string, int)"/>
        /// </summary>
        /// <inheritdoc cref="ImportSheet(string, int)"/>
        /// <returns><see cref="Task{IQueryable{Row}}"/></returns>
        public Task<IQueryable<Row>> ImportSheetAsync(string filePath, int sheetIndex)
        {
            return Task.Factory.StartNew(() => _Import(filePath, sheetIndex));
        }


        /// <summary>
        /// Asynchronously get a worksheet data based on an <see cref="IFormFile"/>
        /// <br/>
        /// <b>Note:</b>
        /// <br/>
        /// This method prefers to be called from an Asp.Net or Asp.Net Core application
        /// </summary>
        /// <typeparam name="T">Type of objects want to get</typeparam>
        /// <param name="file">An <see cref="IFormFile"/> object</param>
        /// <param name="sheetName">Worksheet name</param>
        /// <returns><see cref="Task{List}"/></returns>
        public Task<List<T>> ImportFromIFormFileAsync<T>( IFormFile file , string sheetName ) where T : new()
        {
            return _ImportFromIFormFileAsync<T>( file , sheetName );
        }

        /// <summary>
        /// Asynchronously get a worksheet data based on an <see cref="IFormFile"/>
        /// <br/>
        /// <b>Note:</b>
        /// <br/>
        /// This method prefers to be called from an Asp.Net or Asp.Net Core application
        /// </summary>
        /// <typeparam name="T">Type of objects want to get</typeparam>
        /// <param name="file">An <see cref="IFormFile"/> object</param>
        /// <param name="sheetIndex">Worksheet Index</param>
        /// <returns><see cref="Task{List}"/></returns>
        public Task<List<T>> ImportFromIFormFileAsync<T>( IFormFile file , int sheetIndex ) where T : new()
        {
            return _ImportFromIFormFileAsync<T>( file , sheetIndex );
        }


        /// <summary>
        /// Asynchronously get a worksheet data based on an <see cref="IFormFile"/>
        /// <br/>
        /// <b>Note:</b>
        /// <br/>
        /// This method prefers to be called from an Asp.Net or Asp.Net Core application
        /// </summary>
        /// <typeparam name="T">Type of objects want to get</typeparam>
        /// <typeparam name="TSheetColumnsMap">An implementation of <see cref="ISheetColumnsMap{T}"/></typeparam>
        /// <param name="file">An <see cref="IFormFile"/> object</param>
        /// <param name="sheetName">Worksheet name</param>
        /// <returns><see cref="Task{List}"/></returns>
        public Task<List<T>> ImportFromIFormFileAsync<T, TSheetColumnsMap>( IFormFile file , string sheetName ) where T : class , new() where TSheetColumnsMap : ISheetColumnsMap<T>, new()
        {
            return _ImportFromIFormFileAsync<T, TSheetColumnsMap>( file , sheetName );
        }

        /// <summary>
        /// Asynchronously get a worksheet data based on an <see cref="IFormFile"/>
        /// <br/>
        /// <b>Note:</b>
        /// <br/>
        /// This method prefers to be called from an Asp.Net or Asp.Net Core application
        /// </summary>
        /// <typeparam name="T">Type of objects want to get</typeparam>
        /// <typeparam name="TSheetColumnsMap">An implementation of <see cref="ISheetColumnsMap{T}"/></typeparam>
        /// <param name="file">An <see cref="IFormFile"/> object</param>
        /// <param name="sheetIndex">Worksheet Index</param>
        /// <returns><see cref="Task{List}"/></returns>
        public Task<List<T>> ImportFromIFormFileAsync<T, TSheetColumnsMap>( IFormFile file , int sheetIndex) where T : class , new() where TSheetColumnsMap : ISheetColumnsMap<T>, new()
        {
            return _ImportFromIFormFileAsync<T, TSheetColumnsMap>( file , sheetIndex );
        }

        #endregion

    }
}
