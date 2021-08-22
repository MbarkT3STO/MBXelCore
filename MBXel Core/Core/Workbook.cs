using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using LinqToExcel;
using MBXel_Core.Core.Abstraction;
using MBXel_Core.Core.Units;
using MBXel_Core.Exceptions;
using Spire.Xls;

namespace MBXel_Core.Core
{
    /// <summary>
    /// Represent a workbook
    /// </summary>
    public class Workbook
    {

        #region Private properties

        private readonly Factory.Factory    _factory = new();
        private          Spire.Xls.Workbook _workBook;
        private          List<WorkSheet>    _sheets;
        private          bool               _isFirstSheet = true;

        #endregion

        #region Public properties

        public IWorkbookConfig Configuration { get; set; } = new WorkbookConfig();

        #endregion

        #region Constructors

        /// <summary>
        /// Use this constructor when want to Export a new workbook
        /// </summary>
        /// <param name="path">The file path</param>
        /// <param name="numberOfSheets">Number of sheets to be created by default in the workbook</param>
        /// <param name="extension">The workbook file extension</param>
        /// <param name="version">The workbook file version</param>
        public Workbook(string path, int numberOfSheets = 1, Enums.XLExtension extension = Enums.XLExtension.Xlsx, ExcelVersion version = ExcelVersion.Version2016)
        {
            Configuration.SheetsCount = numberOfSheets;
            Configuration.Extension   = extension;
            Configuration.Version     = version;
            Configuration.Path        = path + (Configuration.Extension is Enums.XLExtension.Xlsx ? ".xlsx" : ".xls");

            Configuration.SheetsCount = Configuration.SheetsCount == 0 ? 1 : Configuration.SheetsCount;
            _workBook                 = _factory.CreateWorkbook(Configuration.SheetsCount);
            _sheets                   = _factory.CreateWorkSheets(Configuration.SheetsCount);
        }
          
        
        /// <summary>
        /// This constructor can be used with all cases
        /// </summary>
        public Workbook(IWorkbookConfig workbookConfig)
        {
            Configuration             = workbookConfig;

            //Configuration.SheetsCount = Configuration.SheetsCount == 0 ? 1 : Configuration.SheetsCount;
            _workBook                 = _factory.CreateWorkbook(Configuration.SheetsCount);
            //_sheets                   = _factory.CreateWorkSheets(Configuration.SheetsCount);
        }


        #endregion

        #region Private methods

        #region Workbook

        private void _SetWorkBookOpenPasswordFromConfig()
        {
            if ( Configuration.Password != null )
            {
                _workBook.OpenPassword = Configuration.Password;
            }
        }
        private void _SetWorkBookOpenVersionFromConfig()
        {
            _workBook.Version = Configuration.Version;
        }

        private void _ApplyWorkbookConfig()
        {
            _SetWorkBookOpenVersionFromConfig();
            _SetWorkBookOpenPasswordFromConfig();
        }
        private int _GetWorkBookSheetsCount() => _workBook.Worksheets.Count;
        private void CreateBackInWorkBookWorkSheet(string sheetName)
        {
            if (_sheets.Count > 1)
            {
                _factory.CreateWorkSheet(ref _workBook, sheetName);
            }
        }
        private void _SetWorkBookPath(string path)
        {
            Configuration.Path = path;
        }

        private void SaveTheWorkBook()
        {
            _ApplyWorkbookConfig();
            _workBook.SaveToFile(Configuration.Path, Configuration.Version);
        }
        private void _ToPDF(string path)
        {
            _workBook.SaveToFile(path, FileFormat.PDF);
        }
        private void _ToXML(string path)
        {
            _workBook.SaveAsXml( path );
        }

        private void _LoadFromFile()
        {
            _workBook = _factory.CreateWorkbook();
            _SetWorkBookOpenPasswordFromConfig();
            _workBook.LoadFromFile(Configuration.Path);

            Configuration.SheetsCount = _GetWorkBookSheetsCount();
            _sheets                   = _factory.CreateWorkSheets(Configuration.SheetsCount);

            for (int i = 0; i < Configuration.SheetsCount; i++)
            {
                if (_sheets[i].Content == null)
                {
                    _SetWorkSheetContent(i);
                }
                _SetTheWorkSheetName(i, _workBook.Worksheets[i].Name);
                _SetWorkSheetContent(i, _workBook.Worksheets[i]);
            }

            _isFirstSheet = false;
        }
        private void _LoadFromStream(Stream stream)
        {
            _workBook = _factory.CreateWorkbook();
            _SetWorkBookOpenPasswordFromConfig();
            _workBook.LoadFromStream( stream );

            Configuration.SheetsCount = _GetWorkBookSheetsCount();
            _sheets                   = _factory.CreateWorkSheets(Configuration.SheetsCount);

            for (int i = 0; i < Configuration.SheetsCount; i++)
            {
                if (_sheets[i].Content == null)
                {
                    _SetWorkSheetContent(i);
                }
                _SetTheWorkSheetName(i, _workBook.Worksheets[i].Name);
                _SetWorkSheetContent(i, _workBook.Worksheets[i]);
            }

            _isFirstSheet = false;
        }

        private void _SetPassword(string password)
        {
            _workBook.UnProtect();
            _workBook.Protect(password);
        }
        private void _Protect(string passwordToOpen, bool isProtectWindow, bool isProtectContent)
        {
            _workBook.Protect(passwordToOpen, isProtectWindow, isProtectContent);
        }
        private void _Unprotect()
        {
            _workBook.UnProtect();
        }
        private void _Unprotect(string bookAndStructurePassword)
        {
            _workBook.UnProtect(bookAndStructurePassword);
        }
        private void _SetAuthor(string author)
        {
            _workBook.DocumentProperties.Author = author;
        }
        private void _SetCompany(string company)
        {
            _workBook.DocumentProperties.Company = company;
        }
        private void _SetVersion(ExcelVersion version)
        {
            _workBook.Version = version;
        }

        #endregion

        #region Worksheet

        private WorkSheet _WorkSheet( int workSheetIndex )
        {
            return _sheets[workSheetIndex];
        }
        private WorkSheet _WorkSheet( string workSheetName )
        {
            return _sheets.FirstOrDefault( x => x.Name == workSheetName );
        }

        private void _UpdateSheetsCount( int newSheetsCount )
        {
            Configuration.SheetsCount = newSheetsCount;
        }
        private void _SetWorkSheetType<T>(int sheetIndex) where  T : class
        {
            _sheets[sheetIndex].Type = typeof(T);
        }
        private void _SetWorkSheetContent(int sheetIndex)
        {
            _sheets[sheetIndex].Content = _workBook.Worksheets[sheetIndex];
        }
        private void _SetWorkSheetContent(int sheetIndex, Worksheet worksheet)
        {
            _sheets[sheetIndex].Content = worksheet;
        }
        private void _SetTheWorkSheetName(int sheetIndex, string sheetName)
        {
            if (sheetName != null && _sheets[sheetIndex].Content != null)
            {
                _sheets[sheetIndex].SetName(sheetName);

            }
            else if (sheetName != null)
            {
                _sheets[sheetIndex].SetName(sheetName);
            }
        }

        private void _CreateNewWorkSheet()
        {
            if (_isFirstSheet)
            {
                _sheets       = _factory.CreateWorkSheets(1);
                _isFirstSheet = false;
            }
            else
            {
                _sheets.Add(new WorkSheet());
            }
        }
        private void _InsertEmptyWorkSheet(string sheetName)
        {
            //if (_sheets == null || _sheets.Count <= 1)
            //{
            //    _sheets = _factory.CreateWorkSheets(1);
            //}
            //else
            //{
            //    ThrowExceptionIfWorkSheetNameIsExist(sheetName);
            //    _sheets.Add(new WorkSheet());
            //}

            if ( _isFirstSheet )
            {
                _sheets       = _factory.CreateWorkSheets(1);
                _isFirstSheet = false;
            }
            else
            {
                _sheets.Add( new WorkSheet() );
            }

            if ( _sheets.Count > 1 )
            {
                _factory.CreateWorkSheet(ref _workBook, sheetName);
            }

            var sheetIndex = _sheets.Count - 1;

            _sheets[sheetIndex].Content = _workBook.Worksheets[sheetIndex];

            Configuration.SheetsCount = _sheets.Count;
            _SetTheWorkSheetName(sheetIndex, sheetName);


            //_sheets.Add(new WorkSheet() { Name = sheetName, Content = _workBook.Worksheets[Configuration.SheetsCount - 1] });
        }   

        private WorkSheet _CreateEmptyWorkSheet<T>(string sheetName) where T : class
        {

            _CreateNewWorkSheet();
            CreateBackInWorkBookWorkSheet( sheetName );

            var sheetIndex = _sheets.Count - 1;

            _UpdateSheetsCount( _sheets.Count );
            _SetWorkSheetContent( sheetIndex );
            _SetWorkSheetType<T>( sheetIndex );

            _SetTheWorkSheetName(sheetIndex, sheetName);

            return _sheets[^1];
        }
        private WorkSheet _CreateEmptyWorkSheet(string sheetName)
        {

            _CreateNewWorkSheet();
            CreateBackInWorkBookWorkSheet(sheetName);

            var sheetIndex = _sheets.Count - 1;

            _UpdateSheetsCount(_sheets.Count);
            _SetWorkSheetContent(sheetIndex);
            _SetWorkSheetType<object>(sheetIndex);

            _SetTheWorkSheetName(sheetIndex, sheetName);

            return _sheets[^1];
        }

        private void CreateWorkSheet<T>(int sheetIndex, List<T> data, string sheetName =null)
        {
            _sheets[sheetIndex].Type    = typeof(T);
            _sheets[sheetIndex].Content = _workBook.Worksheets[sheetIndex];
            var sheet = _sheets[sheetIndex];

            _SetTheWorkSheetName(sheetIndex, sheetName);
            PrepareTheWorkSheetHeaders(sheet);
            PrepareTheWorkSheetData(sheet, data);
            StylingTheWorkSheet(sheet, data.Count);
        }
        private void CreateWorkSheet<T>(int sheetIndex, List<T> data, List<string> columnHeaders, string sheetName=null)
        {
            _sheets[sheetIndex].Type    = typeof(T);
            _sheets[sheetIndex].Content = _workBook.Worksheets[sheetIndex];
            var sheet = _sheets[sheetIndex];

            _SetTheWorkSheetName(sheetIndex, sheetName);
            PrepareTheWorkSheetHeaders(sheet, columnHeaders);
            PrepareTheWorkSheetData(sheet, data);
            StylingTheWorkSheet(sheet, data.Count);
        }

        private void _RemoveWorkSheet(int sheetIndex)
        {
            ThrowExceptionIfWorkSheetIndexNotExist(sheetIndex);

            _sheets.RemoveAt(sheetIndex);
            _workBook.Worksheets[sheetIndex].Remove();
            _UpdateSheetsCount( _sheets.Count );
        }
        private void _RemoveWorkSheet(string sheetName)
        {
            ThrowExceptionIfWorkSheetNameNotExist(sheetName);

            var sheet = _sheets.FirstOrDefault(x => x.Content.Name == sheetName);
            _sheets.Remove(sheet);
            _workBook.Worksheets.Remove(sheetName);
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
            var lastColumnIndex = workSheet.Content.LastColumn;
            var lastRowIndex    = workSheet.Content.LastRow;

            //Columns styling
            workSheet.Content.Range[1, 1, 1, lastColumnIndex].Style.Font.Size = 14;
            workSheet.Content.Range[1, 1, 1, lastColumnIndex].Style.Font.IsBold = true;
            workSheet.Content.Range[1, 1, 1, lastColumnIndex].Style.Font.Color = Color.White;
            workSheet.Content.Range[1, 1, 1, lastColumnIndex].Style.Interior.Color = ColorTranslator.FromHtml("#54a0ff");
            workSheet.Content.Range[1, 1, 1, lastColumnIndex].Style.HorizontalAlignment = HorizontalAlignType.Center;
            workSheet.Content.Range[1, 1, 1, lastColumnIndex].Style.VerticalAlignment = VerticalAlignType.Center;


            //Rows styling
            workSheet.Content.Range[2, 1, lastRowIndex, lastColumnIndex].Style.Font.Size = 14;
            workSheet.Content.Range[2, 1, lastRowIndex, lastColumnIndex].Style.Font.Color = Color.White;
            workSheet.Content.Range[2, 1, lastRowIndex, lastColumnIndex].Style.Interior.Color = ColorTranslator.FromHtml("#2ed573");
            workSheet.Content.Range[2, 1, lastRowIndex, lastColumnIndex].Style.HorizontalAlignment = HorizontalAlignType.Center;
            workSheet.Content.Range[2, 1, lastRowIndex, lastColumnIndex].Style.VerticalAlignment = VerticalAlignType.Center;

            //Other Columns styling
            workSheet.Content.AllocatedRange.AutoFitRows();
            workSheet.Content.AllocatedRange.AutoFitColumns();

            //Other Rows styling
            workSheet.Content.SetRowHeight(1, 30);
        }

        private bool _IsWorksheetContentNull(string sheetName)
        {
            return _sheets.FirstOrDefault(x => x.Name == sheetName && x.Content != null) == null;
        }

        #endregion

        #region Types

        private PropertyInfo[] GetTypeProps<T>()                    => typeof(T).GetProperties();
        private PropertyInfo[] GetTypePropsOfSheet(WorkSheet sheet) => sheet.Type.GetProperties();

        #endregion

        #region Exceptions

        private void ThrowExceptionIfWorkSheetIndexNotExist(int sheetIndex)
        {
            if (_sheets.ElementAtOrDefault(sheetIndex) == null)
                throw new IndexOutOfRangeException("Sheet index was not found");
        }
        private void ThrowExceptionIfWorkSheetNameIsExist(string sheetName)
        {
            if (_sheets.Count > 0 && sheetName != null)
            {
                if (_sheets.FirstOrDefault(x => x.Name == sheetName) != null)
                    throw new SheetNameAlreadyExistException($"A worksheet with the same name ({sheetName}) already exist");
            }
        }
        private void ThrowExceptionIfWorkSheetNameNotExist(string sheetName)
        {
            if (_sheets.FirstOrDefault(x => x.Content.Name == sheetName) == null)
                throw new SheetNameNotExistException($"A worksheet with the name ({sheetName}) doesn't exist");
        }

        #endregion

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
                _SetWorkBookPath(path);

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
            var workbook = new ExcelQueryFactory(Configuration.Path);

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
            var workbook = new ExcelQueryFactory(Configuration.Path);

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

        #region Save Workbook

        public void Save()
        {
            SaveTheWorkBook();
        }

        public Task SaveAsync() => Task.Factory.StartNew(() => Save());

        #endregion

        #region To PDF

        /// <summary>
        /// Convert a workbook to PDF
        /// </summary>
        /// <param name="path">Path to save the PDF in</param>
        public void ToPdf(string path)
        {
            _ToPDF(path);
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="ToPdf"/>
        /// </summary>
        /// <inheritdoc cref="ToPdf"/>
        /// <returns><see cref="Task"/></returns>
        public Task ToPdfAsync(string path) => Task.Factory.StartNew(() => ToPdf(path));

        #endregion

        #region To XML

        /// <summary>
        /// Convert a workbook to Office Open XML
        /// </summary>
        /// <param name="path">Path to save the XML in</param>
        public void ToXml(string path)
        {
            _ToXML( path );
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="ToXml"/>
        /// </summary>
        /// <inheritdoc cref="ToXml"/>
        /// <returns><see cref="Task"/></returns>
        public Task ToXmlAsync( string path ) => Task.Factory.StartNew( () => ToXml( path ) );

        #endregion

        #endregion

        #region Method chaining

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

        #region Create empty worksheet

        public WorkSheet CreateEmptyWorkSheet<T>(string sheetName = null) where  T : class
        {
            return _CreateEmptyWorkSheet<T>( sheetName );
        }
        public WorkSheet CreateEmptyWorkSheet(string sheetName = null)
        {
            return _CreateEmptyWorkSheet( sheetName );
        }

        #endregion

        #region Load

        /// <inheritdoc cref="LoadFromFile"/>
        public Workbook LoadFile( string path = null )
        {
            LoadFromFile( path );
            return this;
        }

        /// <summary>
        /// Load a workbook from a stream
        /// </summary>
        /// <param name="stream">Stream to load workbook from</param>
        public Workbook LoadFromStream(Stream stream)
        {
            _LoadFromStream( stream );
            return this;
        }

        #endregion

        #region password

        /// <summary>
        /// Set a password for the workbook file
        /// </summary>
        /// <param name="password">Custom password</param>
        public Workbook Password( string password )
        {
            _SetPassword( password );
            return this;
        }

        #endregion

        #region Author

        /// <summary>
        /// Set the Author for the workbook
        /// </summary>
        /// <param name="author">Author</param>
        /// <returns><see cref="Workbook"/></returns>
        public Workbook Author(string author)
        {
            _SetAuthor( author );
            return this;
        }

        #endregion
        
        #region Company

        /// <summary>
        /// Set the Company property for the workbook
        /// </summary>
        /// <param name="company">Company</param>
        /// <returns><see cref="Workbook"/></returns>
        public Workbook Company(string company)
        {
            _SetCompany( company );
            return this;
        }

        #endregion    
        
        #region Version

        /// <summary>
        /// Set Version for the workbook
        /// </summary>
        /// <param name="version">Workbook file version</param>
        /// <returns><see cref="Workbook"/></returns>
        public Workbook Version(ExcelVersion version)
        {
            _SetVersion( version );
            return this;
        }

        #endregion

        #region Protect

        /// <summary>
        /// Protect file,also Indicates whether protect workbook window and structure or not
        /// </summary>
        /// <param name="passwordToOpen">password to open file.</param>
        /// <param name="isProtectWindow">Indicates if protect workbook window.</param>
        /// <param name="isProtectContent">Indicates if protect workbook content.</param>
        public Workbook Protect(string passwordToOpen, bool isProtectWindow, bool isProtectContent)
        {
            _Protect(passwordToOpen, isProtectWindow, isProtectContent);
            return this;
        } 
        

        #endregion
        
        #region Unprotect

        /// <summary>
        /// Remove the workbook protection 
        /// </summary>
        public Workbook Unprotect()
        {
            _Unprotect();
            return this;
        } 
        
        /// <summary>
        /// Remove the workbook protection 
        /// </summary>
        public Workbook Unprotect(string bookAndStructurePassword)
        {
            _Unprotect( bookAndStructurePassword );
            return this;
        }

        #endregion

        #region Worksheet

        /// <summary>
        /// Get a specific worksheet from the workbook
        /// </summary>
        /// <param name="workSheetIndex">Worksheet index</param>
        /// <returns><see cref="IWorkSheet"/></returns>
        public WorkSheet Worksheet(int workSheetIndex)
        {
            return _WorkSheet( workSheetIndex );
        }

        /// <summary>
        /// Get a specific worksheet from the workbook
        /// </summary>
        /// <param name="workSheetName">Worksheet name</param>
        /// <returns><see cref="IWorkSheet"/></returns>
        public WorkSheet Worksheet(string workSheetName)
        {
            return _WorkSheet( workSheetName );
        }

        #endregion

        #endregion

    }
}
