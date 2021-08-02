using LinqToExcel;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MBXel_Core.Core
{
    /// <summary>
    /// Import data from an Excel file
    /// </summary>
    public class XLImporter
    {

        #region Private methods

        private IQueryable<Row> _Import(string filePath, string sheetName)
        {
            //Load the workbook
            var workbook = new ExcelQueryFactory(filePath);

            //Collect data from the worksheet
            var result = workbook.Worksheet(sheetName);

            return result;
        }
        
        private IQueryable<Row> _Import(string filePath, int sheetIndex)
        {
            //Load the workbook
            var workbook = new ExcelQueryFactory(filePath);

            //Collect data from the worksheet
            var result = workbook.Worksheet(sheetIndex);

            return result;
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


        #endregion

    }
}
