using MBXel_Core.Exceptions;

using Spire.Xls;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Threading.Tasks;

namespace MBXel_Core.Core
{
    /// <summary>
    /// Export a <see cref="List{T}"/> as an Excel file
    /// </summary>
    class XLExporter
    {
        #region Private properties

        private PropertyInfo[] _properties;
        private Workbook _workBook;
        private Worksheet _workSheet;
        private ExcelVersion _excelVersion;
        private int _TpropertiesCount { get =>  _properties.Length; }

        #endregion

        #region Private methods

        private void ConfigureProps<T>(ExcelVersion version)
        {
            // Get data to be exported properties
            _properties = typeof(T).GetProperties();

            // Prepaire the Workbook and Worksheet
            _workBook = new Workbook();
            _workBook.CreateEmptySheets(1);
            _workSheet = _workBook.Worksheets[0];

            // Set Excel version
            _excelVersion = version;
        }

        private void PrepareTheWorkSheetHeaders()
        {
            // Prepaire column headers
            for (int i = 0; i < _properties.Length; i++)
            {
                _workSheet.Range[1, i + 1].Text = _properties[i].Name;
            }
        }

        private void PrepareTheWorkSheetHeaders(List<string> columnHeaders)
        {
            if (columnHeaders.Count == _properties.Length)
            {
                // Prepaire column headers
                for (int i = 0; i < columnHeaders.Count; i++)
                {
                    _workSheet.Range[1, i + 1].Text = columnHeaders[i];
                }
            }
            else
            {
                throw new HeadersPropertiesNotEqualsToDataPropertiesException();
            }
        }

        private void PrepareTheWorkSheetData<T>(List<T> data)
        {
            // Put data into worksheet
                int rowIndex = 2;

                foreach (T d in data)
                {
                    for (int i = 0; i < _properties.Length; i++)
                    {
                        _workSheet.Range[rowIndex, i + 1].Text = _properties[i].GetValue(d).ToString();
                    }

                    rowIndex++;
                }
        } 
        
        private void StylingTheWorkSheet(int rowsNumber)
        {
            //Columns styling
            _workSheet.Range["A1:BB1"].Style.Font.Size = 14;
            _workSheet.Range["A1:BB1"].Style.Font.IsBold = true;
            _workSheet.Range["A1:BB1"].Style.Font.Color = Color.White;
            _workSheet.Range["A1:BB1"].Style.Interior.Color = ColorTranslator.FromHtml("#54a0ff");
            _workSheet.Range["A1:BB1"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            _workSheet.Range["A1:BB1"].Style.VerticalAlignment = VerticalAlignType.Center;

            //Rows styling
            _workSheet.Range[$"A2:BB{rowsNumber + 1}"].Style.Font.Size = 14;
            _workSheet.Range[$"A2:BB{rowsNumber + 1}"].Style.Font.Color = Color.White;
            _workSheet.Range[$"A2:BB{rowsNumber + 1}"].Style.Interior.Color = ColorTranslator.FromHtml("#2ed573");
            _workSheet.Range[$"A2:BB{rowsNumber + 1}"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            _workSheet.Range[$"A2:BB{rowsNumber + 1}"].Style.VerticalAlignment = VerticalAlignType.Center;

            //Other Columns styling
            _workSheet.AllocatedRange.AutoFitRows();
            _workSheet.AllocatedRange.AutoFitColumns();

            //Other Rows styling
            _workSheet.SetRowHeight(1, 30);
        }

        private void SaveTheWorkBook(string path)
        {
            _workBook.SaveToFile(path, _excelVersion);
        }


        private async void _Export<T>(List<T> data, string path, Enums.XLExtension extension, ExcelVersion version)
        {
            // Prepare the Workbook and Worksheet
            ConfigureProps<T>(version);
        
            // Prepare data
            await Task.Run(() => PrepareTheWorkSheetData(data));

            // Prepare column headers
            await Task.Run(() => PrepareTheWorkSheetHeaders());

            // Styling the worksheet
            await Task.Run(() => StylingTheWorkSheet(data.Count));

            // Save the workbook 
            string wbookPath = path + (extension is Enums.XLExtension.Xlsx ? ".xlsx" : ".xls");
            SaveTheWorkBook(wbookPath);
        }

        private async void _Export<T>(List<T> data, string path, List<string> columnHeaders, Enums.XLExtension extension, ExcelVersion version)
        {
            // Prepare the Workbook and Worksheet
            ConfigureProps<T>(version);

            // Prepare column headers
            await Task.Run(() => PrepareTheWorkSheetHeaders(columnHeaders));

            // Prepare data
            await Task.Run(() => PrepareTheWorkSheetData(data));

            // Styling the worksheet
            await Task.Run(() => StylingTheWorkSheet(data.Count));

            //Save the workbook 
            string wbookPath = path + (extension is Enums.XLExtension.Xlsx ? ".xlsx" : ".xls");
            SaveTheWorkBook(wbookPath);
        }

        #endregion

        #region Routins

        /// <summary>
        /// Export a <see cref="List{T}"/> of data to an excel file
        /// </summary>
        /// <param name="data">Data to be exported</param>
        /// <param name="path">Path to be save in</param>
        /// <param name="extension">Excel file extension</param>
        /// <param name="version">Excel file version</param>
        /// <returns><see cref="bool"/></returns>
        public bool Export<T>(List<T> data, string path, Enums.XLExtension extension = Enums.XLExtension.Xlsx, ExcelVersion version = ExcelVersion.Version2016) where T : class
        {
            _Export(data, path, extension, version);
            return true;
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="Export{T}(List{T}, string, Enums.XLExtension, ExcelVersion)"/>
        /// </summary>
        /// <inheritdoc cref="Export{T}(List{T}, string, Enums.XLExtension, ExcelVersion)"/>
        /// <returns><see cref="Task{TResult}"/></returns>
        public Task<bool> ExportAsync<T>(List<T> data, string path, Enums.XLExtension extension = Enums.XLExtension.Xlsx, ExcelVersion version = ExcelVersion.Version2016) where T : class
        {
            return Task.Factory.StartNew(() =>
            {
                _Export(data, path, extension, version);
                return true;
            });
        }


        /// <summary>
        /// Export a <see cref="List{T}"/> of data to an excel file, with a custom column headers text
        /// </summary>
        /// <param name="data">Data to be exported</param>
        /// <param name="path">Path to be save in</param>
        /// <param name="columnHeaders">Custom column headers text/titles</param>
        /// <param name="extension">Excel file extension</param>
        /// <param name="version">Excel file version</param>
        /// <returns><see cref="bool"/></returns>
        public bool Export<T>(List<T> data, string path, List<string> columnHeaders, Enums.XLExtension extension = Enums.XLExtension.Xlsx, ExcelVersion version = ExcelVersion.Version2016) where T : class
        {
            _Export(data, path, columnHeaders, extension, version);
            return true;
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="Export{T}(List{T}, string, List{string}, Enums.XLExtension, ExcelVersion)"/>
        /// </summary>
        /// <inheritdoc cref="Export{T}(List{T}, string, List{string}, Enums.XLExtension, ExcelVersion)"/>
        /// <returns><see cref="Task{TResult}"/></returns>
        public Task<bool> ExportAsync<T>(List<T> data, string path, List<string> columnHeaders, Enums.XLExtension extension = Enums.XLExtension.Xlsx, ExcelVersion version = ExcelVersion.Version2016) where T : class
        {
            return Task.Factory.StartNew(() =>
            {
                _Export(data, path, columnHeaders, extension, version);
                return true;
            });
        }


        #endregion
    }
}
