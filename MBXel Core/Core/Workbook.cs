using MBXel_Core.Exceptions;

using Spire.Xls;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace MBXel_Core.Core.Units
{
    /// <summary>
    /// Represent a workbook
    /// </summary>
    class Workbook
    {
        #region Private properties

        private Factory.Factory     _factory = new();
        private Spire.Xls.Workbook  _workBook;
        private List<WorkSheet>     _sheets;

        #endregion

        #region Public properties

        public string               Path            { get; private set;  }
        public int                  SheetsCount     { get; private set;  }
        public Enums.XLExtension    Extension       { get; private set; }
        public ExcelVersion         Version         { get; private set; }

        #endregion


        public Workbook(string path, int numberOfSheets=1, Enums.XLExtension extension = Enums.XLExtension.Xlsx, ExcelVersion version = ExcelVersion.Version2016)
        {
            SheetsCount = numberOfSheets;
            Extension = extension;
            Version = version;
            Path = path + (Extension is Enums.XLExtension.Xlsx ? ".xlsx" : ".xls");

            _workBook = _factory.CreateWorkbook(SheetsCount);
            _sheets = _factory.CreateWorkSheets(SheetsCount);
        }


        #region Private methods

        private PropertyInfo[] GetTypeProps<T>() => typeof(T).GetProperties();
        private PropertyInfo[] GetTypePropsOfSheet(WorkSheet sheet) => sheet.Type.GetProperties();

        private void SetTheWorkSheetName(int sheetIndex, string sheetName)
        {
            if (sheetName != null)
            {
                _sheets[sheetIndex].SetName(sheetName);
            }
        }

        private void PrepareTheWorkSheetHeaders(WorkSheet workSheet)
        {
          var properties = GetTypePropsOfSheet(workSheet);

          // Prepaire column headers
          for (int i = 0; i < properties.Length; i++)
          {
              workSheet.Content.Range[1, i + 1].Text = properties[i].Name;
          }
        }

        private void PrepareTheWorkSheetHeaders(WorkSheet workSheet, List<string> columnHeaders)
        {
            var properties = GetTypePropsOfSheet(workSheet);

            if (columnHeaders.Count == properties.Length)
            {
                // Prepaire column headers
                for (int i = 0; i < columnHeaders.Count; i++)
                {
                    workSheet.Content.Range[1, i + 1].Text = columnHeaders[i];
                }
            }
            else
            {
                throw new HeadersPropertiesNotEqualsToDataPropertiesException();
            }
        }

        private void PrepareTheWorkSheetData<T>(WorkSheet workSheet, List<T> data)
        {
            var properties = GetTypePropsOfSheet(workSheet);

            // Put data into worksheet
            int rowIndex = 2;

            foreach (T d in data)
            {
                for (int i = 0; i < properties.Length; i++)
                {
                    workSheet.Content.Range[rowIndex, i + 1].Text = properties[i].GetValue(d).ToString();
                }

                rowIndex++;
            }
        }

        private void StylingTheWorkSheet(WorkSheet workSheet, int rowsNumber)
            {
                //Columns styling
                workSheet.Content.Range["A1:BB1"].Style.Font.Size = 14;
                workSheet.Content.Range["A1:BB1"].Style.Font.IsBold = true;
                workSheet.Content.Range["A1:BB1"].Style.Font.Color = Color.White;
                workSheet.Content.Range["A1:BB1"].Style.Interior.Color = ColorTranslator.FromHtml("#54a0ff");
                workSheet.Content.Range["A1:BB1"].Style.HorizontalAlignment = HorizontalAlignType.Center;
                workSheet.Content.Range["A1:BB1"].Style.VerticalAlignment = VerticalAlignType.Center;

                //Rows styling
                workSheet.Content.Range[$"A2:BB{rowsNumber + 1}"].Style.Font.Size = 14;
                workSheet.Content.Range[$"A2:BB{rowsNumber + 1}"].Style.Font.Color = Color.White;
                workSheet.Content.Range[$"A2:BB{rowsNumber + 1}"].Style.Interior.Color = ColorTranslator.FromHtml("#2ed573");
                workSheet.Content.Range[$"A2:BB{rowsNumber + 1}"].Style.HorizontalAlignment = HorizontalAlignType.Center;
                workSheet.Content.Range[$"A2:BB{rowsNumber + 1}"].Style.VerticalAlignment = VerticalAlignType.Center;

                //Other Columns styling
                workSheet.Content.AllocatedRange.AutoFitRows();
                workSheet.Content.AllocatedRange.AutoFitColumns();

                //Other Rows styling
                workSheet.Content.SetRowHeight(1, 30);
            }

        private void SaveTheWorkBook()
        {
            _workBook.SaveToFile(Path, Version);
        }


        private void CheckIfWorkSheetIndexIsExist(int sheetIndex)
        {
            if (_sheets.ElementAtOrDefault(sheetIndex) == null)
                throw new IndexOutOfRangeException("Sheet index was not found");
        }

        private void CreateWorkSheet<T>(int sheetIndex, List<T> data, string sheetName=null)
        {
            _sheets[sheetIndex].Type = typeof(T);
            _sheets[sheetIndex].Content = _workBook.Worksheets[sheetIndex];
            var sheet = _sheets[sheetIndex];

            SetTheWorkSheetName(sheetIndex, sheetName);
            PrepareTheWorkSheetHeaders(sheet);
            PrepareTheWorkSheetData(sheet, data);
            StylingTheWorkSheet(sheet, data.Count);
        }

        private void CreateWorkSheet<T>(int sheetIndex, List<T> data, List<string> columnHeaders, string sheetName=null)
        {
            _sheets[sheetIndex].Type = typeof(T);
            _sheets[sheetIndex].Content = _workBook.Worksheets[sheetIndex];
            var sheet = _sheets[sheetIndex];

            SetTheWorkSheetName(sheetIndex, sheetName);
            PrepareTheWorkSheetHeaders(sheet, columnHeaders);
            PrepareTheWorkSheetData(sheet, data);
            StylingTheWorkSheet(sheet, data.Count);
        }

        private void _SetWorkbookPassword(string password)
        {
            _workBook.Protect(password, true, true);
        }

        #endregion

        #region Public methods

        #region Create Sheet

        /// <summary>
        /// Create and fill a specific Worksheet
        /// </summary>
        /// <typeparam name="T">Type of data to be stored in the worksheet</typeparam>
        /// <param name="sheetIndex">Worksheet index</param>
        /// <param name="data">Data to be stored in the worksheet</param>
        public void CreateSheet<T>(int sheetIndex, List<T> data)
        {
            CheckIfWorkSheetIndexIsExist(sheetIndex);
            CreateWorkSheet<T>(sheetIndex, data);
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="CreateSheet{T}(int, List{T})"/>
        /// </summary>
        /// <inheritdoc cref="CreateSheet{T}(int, List{T})"/>
        /// <returns><see cref="Task"/></returns>
        public Task CreateSheetAsync<T>(int sheetIndex, List<T> data)
        {
            return Task.Factory.StartNew(() => CreateSheet<T>(sheetIndex, data));
        }


        /// <summary>
        /// Create and fill a specific Worksheet with a custom name
        /// </summary>
        /// <typeparam name="T">Type of data to be stored in the worksheet</typeparam>
        /// <param name="sheetIndex">Worksheet index</param>
        /// <param name="sheetName">Worksheet name</param>
        /// <param name="data">Data to be stored in the worksheet</param>
        public void CreateSheet<T>(int sheetIndex, string sheetName, List<T> data)
        {
            CheckIfWorkSheetIndexIsExist(sheetIndex);
            CreateWorkSheet<T>(sheetIndex, data, sheetName);
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="CreateSheet{T}(int, string, List{T})"/>
        /// </summary>
        /// <inheritdoc cref="CreateSheet{T}(int, string, List{T})"/>
        /// <returns><see cref="Task"/></returns>
        public Task CreateSheetAsync<T>(int sheetIndex, string sheetName, List<T> data)
        {
            return Task.Factory.StartNew(() => CreateSheet<T>(sheetIndex, sheetName, data));
        }


        /// <summary>
        /// Create and fill a specific Worksheet with a custom name and headers
        /// </summary>
        /// <typeparam name="T">Type of data to be stored in the worksheet</typeparam>
        /// <param name="sheetIndex">Worksheet index</param>
        /// <param name="data">Data to be stored in the worksheet</param>
        /// <param name="columnHeaders">Custom column headers text/titles</param>
        /// <param name="sheetName">Worksheet name</param>
        public void CreateSheet<T>(int sheetIndex, List<T> data, List<string> columnHeaders = null, string sheetName = null)
        {
            CheckIfWorkSheetIndexIsExist(sheetIndex);

            if (columnHeaders == null)
                CreateWorkSheet<T>(sheetIndex, data, sheetName);
            else
                CreateWorkSheet<T>(sheetIndex, data, columnHeaders, sheetName);
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="CreateSheet{T}(int, List{T}, List{string}, string)"/>
        /// </summary>
        /// <inheritdoc cref="CreateSheet{T}(int, List{T}, List{string}, string)"/>
        /// <returns><see cref="Task"/></returns>
        public Task CreateSheetAsync<T>(int sheetIndex, List<T> data, List<string> columnHeaders = null, string sheetName = null)
        {
            return Task.Factory.StartNew(() => CreateSheet<T>(sheetIndex, data, columnHeaders, sheetName));
        }

        #endregion

        #region Workbook Protection

        /// <summary>
        /// Set a password for the workbook file
        /// </summary>
        /// <param name="password">Custom password</param>
        public void SetPassword(string password)
        {
            _SetWorkbookPassword(password);
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="SetPassword(string)"/>
        /// </summary>
        /// <inheritdoc cref="SetPassword(string)"/>
        /// <returns><see cref="Task"/></returns>
        public Task SetPasswordAsync(string password)
        {
            return Task.Factory.StartNew(() => _SetWorkbookPassword(password));
        }

        #endregion

        #region Save

        public void Save()
        {
            SaveTheWorkBook();
        }

        public Task SaveAsync() => Task.Factory.StartNew(() => Save());

        #endregion

        #endregion

    }
}
