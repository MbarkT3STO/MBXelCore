using LinqToExcel;

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
    public class Workbook
    {

        #region Private properties

        private Factory.Factory     _factory = new();
        private Spire.Xls.Workbook  _workBook;
        private List<WorkSheet> _sheets;

        #endregion

        #region Public properties

        public string               Path            { get; private set;  }
        public int                  SheetsCount     { get; private set;  }
        public Enums.XLExtension    Extension       { get; private set; }
        public ExcelVersion         Version         { get; private set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Use this constructor when want to Import from a saved workbook file
        /// </summary>
        /// <param name="path">The file path</param>
        public Workbook(string path)
        {
            SetWorkBookPath(path);
        }

        /// <summary>
        /// Use this constructor when want to Export a new workbook
        /// </summary>
        /// <param name="path">The file path</param>
        /// <param name="numberOfSheets">Number of sheets to be created by default in the workbook</param>
        /// <param name="extension">The workbook file extension</param>
        /// <param name="version">The workbook file version</param>
        public Workbook(string path, int numberOfSheets = 1, Enums.XLExtension extension = Enums.XLExtension.Xlsx, ExcelVersion version = ExcelVersion.Version2016)
        {
            SheetsCount = numberOfSheets;
            Extension = extension;
            Version = version;
            Path = path + (Extension is Enums.XLExtension.Xlsx ? ".xlsx" : ".xls");

            _workBook = _factory.CreateWorkbook(SheetsCount);
            _sheets = _factory.CreateWorkSheets(SheetsCount);
        }


        #endregion

        #region Private methods

        private PropertyInfo[] GetTypeProps<T>() => typeof(T).GetProperties();
        private PropertyInfo[] GetTypePropsOfSheet(WorkSheet sheet) => sheet.Type.GetProperties();

        private void SetWorkBookPath(string path)
        {
            Path = path;
        }

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

          // Prepare column headers
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
                // Prepare column headers
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

        private void _InsertEmptyWorkSheet(string sheetName)
        {
            ThrowExceptionIfWorkSheetNameIsExist(sheetName);

            _factory.CreateWorkSheet(ref _workBook, sheetName);
            
            SheetsCount += 1;

            _sheets.Add(new WorkSheet() { Content = _workBook.Worksheets[SheetsCount - 1] });
        }

        private void SaveTheWorkBook()
        {
            _workBook.SaveToFile(Path, Version);
        }


        private void ThrowExceptionIfWorkSheetIndexNotExist(int sheetIndex)
        {
            if (_sheets.ElementAtOrDefault(sheetIndex) == null)
                throw new IndexOutOfRangeException("Sheet index was not found");
        }

        private void ThrowExceptionIfWorkSheetNameIsExist(string sheetName)
        {
            if (_sheets.FirstOrDefault( x => x.Content.Name == sheetName ) != null)
                throw new SheetNameAlreadyExistException($"A worksheet with the same name ({sheetName}) already exist");
        }
        
        private void ThrowExceptionIfWorkSheetNameNotExist(string sheetName)
        {
            if (_sheets.FirstOrDefault( x => x.Content.Name == sheetName ) == null)
                throw new SheetNameNotExistException($"A worksheet with the name ({sheetName}) doesn't exist");
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

        private void _RemoveWorkSheet(int sheetIndex)
        {
            ThrowExceptionIfWorkSheetIndexNotExist(sheetIndex);

            _sheets.RemoveAt(sheetIndex);
            _workBook.Worksheets[sheetIndex].Remove();
        }
        
        private void _RemoveWorkSheet(string sheetName)
        {
            ThrowExceptionIfWorkSheetNameNotExist(sheetName);

            var sheet = _sheets.FirstOrDefault(x => x.Content.Name == sheetName);
            _sheets.Remove(sheet);
            _workBook.Worksheets.Remove(sheetName);
        }


        private void _SetWorkbookPassword(string password)
        {
            _workBook.Protect(password, true, true);
        }

        private void _LoadFromFile(Enums.XLExtension extension = Enums.XLExtension.Xlsx, ExcelVersion version = ExcelVersion.Version2016)
        {
            Extension = extension;
            Version = version;

            _workBook = _factory.CreateWorkbook();
            _workBook.LoadFromFile(Path);

            SheetsCount = _workBook.Worksheets.Count;
            _sheets = _factory.CreateWorkSheets(SheetsCount);

            for (int i = 0; i < SheetsCount; i++)
            {
                _sheets[i].Content = _workBook.Worksheets[i];
                _sheets[i].Name = _workBook.Worksheets[i].Name;
            }
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
        public void BuildSheet<T>(int sheetIndex, List<T> data)
        {
            ThrowExceptionIfWorkSheetIndexNotExist(sheetIndex);
            CreateWorkSheet<T>(sheetIndex, data);
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="BuildSheet{T}(int, List{T})"/>
        /// </summary>
        /// <inheritdoc cref="BuildSheet{T}(int, List{T})"/>
        /// <returns><see cref="Task"/></returns>
        public Task BuildSheetAsync<T>(int sheetIndex, List<T> data)
        {
            return Task.Factory.StartNew(() => BuildSheet<T>(sheetIndex, data));
        }


        /// <summary>
        /// Create and fill a specific Worksheet with a custom name
        /// </summary>
        /// <typeparam name="T">Type of data to be stored in the worksheet</typeparam>
        /// <param name="sheetIndex">Worksheet index</param>
        /// <param name="sheetName">Worksheet name</param>
        /// <param name="data">Data to be stored in the worksheet</param>
        public void BuildSheet<T>(int sheetIndex, string sheetName, List<T> data)
        {
            ThrowExceptionIfWorkSheetIndexNotExist(sheetIndex);
            CreateWorkSheet<T>(sheetIndex, data, sheetName);
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="BuildSheet{T}(int, string, List{T})"/>
        /// </summary>
        /// <inheritdoc cref="BuildSheet{T}(int, string, List{T})"/>
        /// <returns><see cref="Task"/></returns>
        public Task BuildSheetAsync<T>(int sheetIndex, string sheetName, List<T> data)
        {
            return Task.Factory.StartNew(() => BuildSheet<T>(sheetIndex, sheetName, data));
        }


        /// <summary>
        /// Create and fill a specific Worksheet with a custom name and headers
        /// </summary>
        /// <typeparam name="T">Type of data to be stored in the worksheet</typeparam>
        /// <param name="sheetIndex">Worksheet index</param>
        /// <param name="data">Data to be stored in the worksheet</param>
        /// <param name="columnHeaders">Custom column headers text/titles</param>
        /// <param name="sheetName">Worksheet name</param>
        public void BuildSheet<T>(int sheetIndex, List<T> data, List<string> columnHeaders = null, string sheetName = null)
        {
            ThrowExceptionIfWorkSheetIndexNotExist(sheetIndex);

            if (columnHeaders == null)
                CreateWorkSheet<T>(sheetIndex, data, sheetName);
            else
                CreateWorkSheet<T>(sheetIndex, data, columnHeaders, sheetName);
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="BuildSheet{T}(int, List{T}, List{string}, string)"/>
        /// </summary>
        /// <inheritdoc cref="BuildSheet{T}(int, List{T}, List{string}, string)"/>
        /// <returns><see cref="Task"/></returns>
        public Task BuildSheetAsync<T>(int sheetIndex, List<T> data, List<string> columnHeaders = null, string sheetName = null)
        {
            return Task.Factory.StartNew(() => BuildSheet<T>(sheetIndex, data, columnHeaders, sheetName));
        }

        #endregion

        #region Remove Sheet

        /// <summary>
        /// Remove an exist worksheet
        /// </summary>
        /// <param name="sheetIndex">Worksheet index</param>
        /// <exception cref="IndexOutOfRangeException"></exception>
        public void RemoveSheet(int sheetIndex)
        {
            _RemoveWorkSheet(sheetIndex);
        } 
        
        /// <summary>
        /// Remove an exist worksheet
        /// </summary>
        /// <param name="sheetName">Worksheet name</param>
        /// <exception cref="IndexOutOfRangeException"></exception>
        public void RemoveSheet(string sheetName)
        {
            _RemoveWorkSheet(sheetName);
        }

        #endregion

        #region Insert new Worksheet

        /// <summary>
        /// Insert a new empty worksheet
        /// </summary>
        /// <param name="sheetName">The new worksheet name</param>
        public void InsertEmptySheet(string sheetName=null)
        {
            _InsertEmptyWorkSheet(sheetName);
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="InsertEmptySheet(string)"/>
        /// </summary>
        /// <returns><see cref="Task"/></returns>
        /// <inheritdoc cref="InsertEmptySheet(string)"/>
        public Task InsertEmptySheetAsync(string sheetName = null)
        {
            return Task.Factory.StartNew( () => InsertEmptySheet(sheetName));
        }

        #endregion

        #region Load Workbook

        /// <summary>
        /// Load an excel file and import its data
        /// </summary>
        /// <param name="path">File Name/Path, <b>can be ignored if a path provided in the constructor</b></param>
        public void LoadFromFile(string path = null)
        {
            if (path != null)
                SetWorkBookPath(path);

            _LoadFromFile();
        }

        #endregion

        #region Import from worksheet

        /// <summary>
        /// Get a specific worksheet data available for querying
        /// </summary>
        /// <param name="sheetIndex">Worksheet index</param>
        /// <returns><see cref="IQueryable{Row}"/></returns>
        public IQueryable<Row> GetSheetAsQueryable(int sheetIndex)
        {
            //Load the workbook
            var workbook = new ExcelQueryFactory(Path);

            //Collect data from the worksheet
            var result = workbook.Worksheet(sheetIndex);

            return result;
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="GetSheetAsQueryable(int)"/>
        /// </summary>
        /// <inheritdoc cref="GetSheetAsQueryable(int)"/>
        /// <returns><see cref="Task{IQueryable{Row}}"/></returns>
        public Task<IQueryable<Row>> GetSheetAsQueryableAsync(int sheetIndex)
        {
            return Task.Factory.StartNew( () => GetSheetAsQueryable(sheetIndex) );
        }


        /// <summary>
        /// Get a specific worksheet data available for querying
        /// </summary>
        /// <param name="sheetName">Worksheet name</param>
        /// <returns><see cref="IQueryable{Row}"/></returns>
        public IQueryable<Row> GetSheetAsQueryable(string sheetName)
        {
            //Load the workbook
            var workbook = new ExcelQueryFactory(Path);

            //Collect data from the worksheet
            var result = workbook.Worksheet(sheetName); ;

            return result;
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="GetSheetAsQueryable(string)"/>
        /// </summary>
        /// <inheritdoc cref="GetSheetAsQueryable(string)"/>
        /// <returns><see cref="Task{IQueryable{Row}}"/></returns>
        public Task<IQueryable<Row>> GetSheetAsQueryableAsync(string sheetName)
        {
            return Task.Factory.StartNew(() => GetSheetAsQueryable(sheetName));
        }


        /// <summary>
        /// Querying directly on the worksheet
        /// </summary>
        /// <param name="sheetIndex">Worksheet index</param>
        /// <param name="selector">Custom <see cref="Func{T, TResult}"/> selector</param>
        /// <returns><see cref="IQueryable{Row}"/></returns>
        public List<T> Select<T>(int sheetIndex, Func<Row, T> selector) where T : class
        {
            // Get worksheet data as Queryble
            var sheetAsQueryable = GetSheetAsQueryable(sheetIndex);

            // Apply the selector
            var result = sheetAsQueryable.Select(selector).AsParallel().ToList();

            return result;
        }

        /// <summary>
        /// Aynchronously, <inheritdoc cref="Select{T}(int, Func{Row, T})"/>
        /// </summary>
        /// <returns><see cref="Task{List{T}}"/></returns>
        /// <inheritdoc cref="Select{T}(int, Func{Row, T})"/>
        public Task<List<T>> SelectAsync<T>(int sheetIndex, Func<Row, T> selector) where T : class
        {
            return Task.Factory.StartNew( () => Select<T>(sheetIndex, selector) );
        }


        /// <summary>
        /// Querying directly on the worksheet
        /// </summary>
        /// <param name="sheetName">Worksheet name</param>
        /// <param name="selector">Custom <see cref="Func{T, TResult}"/> selector</param>
        /// <returns><see cref="IQueryable{Row}"/></returns>
        public List<T> Select<T>(string sheetName, Func<Row, T> selector) where T : class
        {
            // Get worksheet data as Queryble
            var sheetAsQueryable = GetSheetAsQueryable(sheetName);

            // Apply the selector
            var result = sheetAsQueryable.Select(selector).AsParallel().ToList();

            return result;
        }

        /// <summary>
        /// Aynchronously, <inheritdoc cref="Select{T}(string, Func{Row, T})"/>
        /// </summary>
        /// <returns><see cref="Task{List{T}}"/></returns>
        /// <inheritdoc cref="Select{T}(string, Func{Row, T})"/>
        public Task<List<T>> SelectAsync<T>(string sheetName, Func<Row, T> selector) where T : class
        {
            return Task.Factory.StartNew(() => Select<T>(sheetName, selector));
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

        #region Save Workbook

        public void Save()
        {
            SaveTheWorkBook();
        }

        public Task SaveAsync() => Task.Factory.StartNew(() => Save());

        #endregion

        #endregion

        #region Chaining methods

        #region Build Worksheet

        /// <inheritdoc cref="BuildSheet{T}(int,List{T})"/>
        public Workbook BuildWorkSheet<T>(int sheetIndex , List<T> data )
        {
            BuildSheet<T>( sheetIndex , data );
            return this;
        }

        /// <inheritdoc cref="BuildSheet{T}(int,string,List{T})"/>
        public Workbook BuildWorkSheet<T>( int sheetIndex , string sheetName , List<T> data )
        {
            BuildSheet<T>(sheetIndex, sheetName, data);
            return this;
        }

        /// <inheritdoc cref="BuildSheet{T}(int,List{T},List{string},string)"/>
        public Workbook BuildWorkSheet<T>( int sheetIndex , List<T> data , List<string> columnHeaders = null , string sheetName = null )
        {
            BuildSheet<T>( sheetIndex , data , columnHeaders , sheetName );
            return this;
        }

        #endregion

        #region Remove worksheet

        ///<inheritdoc cref="RemoveSheet(int)"/>
        public Workbook RemoveWorkSheet( int sheetIndex )
        {
            RemoveSheet(sheetIndex);
            return this;
        } 
        
        ///<inheritdoc cref="RemoveSheet(string)"/>
        public Workbook RemoveWorkSheet(string sheetName)
        {
            RemoveSheet( sheetName );
            return this;
        }

        #endregion

        #region Insert new worksheet

        /// <inheritdoc cref="InsertEmptySheet"/>
        public Workbook InsertEmptyWorkSheet( string sheetName = null )
        {
            InsertEmptySheet(sheetName);
            return this;
        }

        #endregion

        #region Load Workbook

        /// <inheritdoc cref="LoadFromFile"/>
        public Workbook LoadFile( string path = null )
        {
            LoadFromFile( path );
            return this;
        }

        #endregion

        #region Set password

        /// <inheritdoc cref="SetPassword"/>
        public Workbook Protect( string password )
        {
            SetPassword( password );
            return this;
        }

        #endregion

        #endregion

    }
}
