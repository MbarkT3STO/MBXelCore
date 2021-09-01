using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;
using MBXel_Core.Core.Abstraction;
using MBXel_Core.Core.Units;
using MBXel_Core.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using Spire.Pdf.HtmlConverter.Qt;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace MBXel_Core.Extensions
{
    public static class WorkSheetExtensions
    {

        #region Private methods

        private static bool IsOneOfObjectPropertiesNull<T>(T obj)
        {
            var propertiesOfT = typeof(T).GetProperties();

            foreach (var prop in propertiesOfT)
            {
                var value = obj.GetType().GetProperty(prop.Name).GetValue(obj, null);
                if (value == null)
                {
                    return true;
                }
            }

            return false;
        }
        private static bool IsOneOfObjectPropertiesNull<T>(T obj, Expression<Func<T, object>> propertiesToBeChecked)
        {
            var propertiesOfT = GetUsedProperties( propertiesToBeChecked );

            foreach (var prop in propertiesOfT)
            {
                var value = obj.GetType().GetProperty(prop.Name).GetValue(obj, null);
                if (value == null)
                {
                    return true;
                }
            }

            return false;
        }
        private static void ChangeTypeOfWorkSheet<T>( ref WorkSheet workSheet, List<T> data )
        {
            workSheet.Type = typeof( T );
        }
        private static PropertyInfo[] GetTypePropsOfSheet(WorkSheet sheet) => sheet.Type.GetProperties();
        private static Dictionary<string, int> GetSheetHeadersIndexes<T>(WorkSheet worksheet, int headersRowIndex)
        {
            var result = new Dictionary<string, int>();

            var propsCountOfT = typeof(T).GetProperties().Length;

            var headers    = new List<string>();
            var headersRow = worksheet.Content.Rows[headersRowIndex];

            for (int i = 0; i < propsCountOfT; i++)
            { 
                var currentTitle = headersRow.Cells[i].Text;

                if (currentTitle.Trim().Length != 0)
                {
                    result.Add(currentTitle, i);
                }
            }

            return result;
        }

        private static void PrepareTheWorkSheetHeaders(ref WorkSheet workSheet)
        {
            var properties = GetTypePropsOfSheet(workSheet);

            // Prepare column headers
            for (int i = 0; i < properties.Length; i++)
            {
                workSheet.Content.Range[1, i + 1].Text = properties[i].Name;
            }
        }
        private static void PrepareTheWorkSheetHeaders<T, TSheetColumnsMap>(ref WorkSheet workSheet) where T : class where TSheetColumnsMap : ISheetColumnsMap<T> , new()
        {
            var properties = GetTypePropsOfSheet(workSheet);

            var headers = new TSheetColumnsMap().CreateMap();

            // Prepare column headers
            for (int i = 0; i < properties.Length; i++)
            {
                workSheet.Content.Range[1 , i + 1].Text = headers[properties[i].Name];
            }
        }
        private static void PrepareTheWorkSheetHeaders(ref WorkSheet workSheet, IReadOnlyList<string> columnHeaders)
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
        private static void PrepareTheWorkSheetData<T>(ref WorkSheet workSheet, List<T> data)
        {
            var properties = GetTypePropsOfSheet( workSheet );

            // Put data into worksheet
            int rowIndex = 2;

            foreach (T obj in data)
            {
                for (int i = 0; i < properties.Length; i++)
                {
                    if (properties[i].GetValue(obj) != null)
                    {
                        workSheet.Content.Range[rowIndex , i + 1].Text = properties[i].GetValue( obj ).ToString();
                    }
                }

                rowIndex++;
            }
        }

        private static void StylingTheWorkSheet(ref WorkSheet workSheet, int rowsNumber)
        {
            var lastColumnIndex = workSheet.Content.LastColumn;
            var lastRowIndex    = workSheet.Content.LastRow;

            //Columns styling
            workSheet.Content.Range[1 , 1 , 1 , lastColumnIndex].Style.Font.Size = 14;
            workSheet.Content.Range[1 , 1 , 1 , lastColumnIndex].Style.Font.IsBold = true;
            workSheet.Content.Range[1 , 1 , 1 , lastColumnIndex].Style.Font.Color = Color.White;
            workSheet.Content.Range[1 , 1 , 1 , lastColumnIndex].Style.Interior.Color = ColorTranslator.FromHtml( "#54a0ff" );
            workSheet.Content.Range[1 , 1 , 1 , lastColumnIndex].Style.HorizontalAlignment = HorizontalAlignType.Center;
            workSheet.Content.Range[1 , 1 , 1 , lastColumnIndex].Style.VerticalAlignment = VerticalAlignType.Center;
            

            //Rows styling
            workSheet.Content.Range[2 , 1 , lastRowIndex , lastColumnIndex].Style.Font.Size = 14;
            workSheet.Content.Range[2 , 1 , lastRowIndex , lastColumnIndex].Style.Font.Color = Color.White;
            workSheet.Content.Range[2 , 1 , lastRowIndex , lastColumnIndex].Style.Interior.Color = ColorTranslator.FromHtml( "#2ed573" );
            workSheet.Content.Range[2 , 1 , lastRowIndex , lastColumnIndex].Style.HorizontalAlignment = HorizontalAlignType.Center;
            workSheet.Content.Range[2 , 1 , lastRowIndex , lastColumnIndex].Style.VerticalAlignment = VerticalAlignType.Center;

            //Other Columns styling
            workSheet.Content.AllocatedRange.AutoFitRows();
            workSheet.Content.AllocatedRange.AutoFitColumns();

            //Other Rows styling
            workSheet.Content.SetRowHeight(1, 30);
        }
        private static void StylingTheHeader(ref WorkSheet workSheet, string backColor, string fontColor, int headerRowIndex = 0)
        {
            var lastColumnIndex = workSheet.Content.LastColumn;

            // Header styling
            workSheet.Content.Range[headerRowIndex + 1 , 1 , headerRowIndex + 1 , lastColumnIndex].Style.Font.Color = ColorTranslator.FromHtml( fontColor );
            workSheet.Content.Range[headerRowIndex + 1 , 1 , headerRowIndex + 1 , lastColumnIndex].Style.Interior.Color = ColorTranslator.FromHtml( backColor );
        }
        private static void StylingTheBody(ref WorkSheet workSheet, string backColor, string fontColor, int bodyStartRowIndex = 0)
        {
            var lastColumnIndex = workSheet.Content.LastColumn;
            var lastRowIndex    = workSheet.Content.LastRow;

            // Body styling
            workSheet.Content.Range[bodyStartRowIndex + 1 , 1 , lastRowIndex , lastColumnIndex].Style.Font.Color = ColorTranslator.FromHtml( fontColor );
            workSheet.Content.Range[bodyStartRowIndex + 1 , 1 , lastRowIndex , lastColumnIndex].Style.Interior.Color = ColorTranslator.FromHtml( backColor );
        }

        private static List<PropertyInfo> GetUsedProperties<T>(Expression<Func<T,object>> usedPropertiesExpression )
        {
            var props = ( from prop in usedPropertiesExpression.Body.Type.GetProperties() select prop ).ToList();

            return props;
        }
        private static List<string> GetCellRangeValues(IEnumerable<CellRange> cells)
        {
            var result = new List<string>();
            foreach ( var cell in cells )
            {
                result.Add(cell.Value);
            }

            return result;
        }
        private static Dictionary<string, int> GetHeaderTitlesWithIndexes(WorkSheet workSheet, int headerRowIndex)
        {
            var result          = new Dictionary<string , int>();
            var lastColumnIndex = workSheet.Content.LastColumn;
            var headerRow       = workSheet.Content.Rows[headerRowIndex];
            var headerTitles    = GetCellRangeValues( headerRow.Cells[..lastColumnIndex] );

            for ( int i = 0 ; i < lastColumnIndex ; i ++ )
            {
                var currentCell      = headerRow.Cells[i];
                var currentCellValue = currentCell.Value;
                var currentCellIndex = headerTitles.IndexOf( currentCellValue );
                result.Add( currentCellValue , currentCellIndex );
            }

            return result;
        }

        private static List<T> Select<T>(WorkSheet workSheet, int headerRowIndex, bool ignoreObjectIfOnePropertyHasNoValue) where  T : class, new()
        {
            var propertiesOfT             = typeof(T).GetProperties();
            var propertiesOfTAsList       = propertiesOfT.ToList();
            var cells                     = workSheet.Content.Cells;
            var lastUsedRowInTheWorkSheet = workSheet.Content.LastRow;
            var result                    = new List<T>();

            for (var i = headerRowIndex + 2; i <= lastUsedRowInTheWorkSheet; i++)
            {
                var row = cells[i];
                var obj = new T();

                foreach (var prop in propertiesOfT)
                {
                    var currentCellValue = row[i, propertiesOfTAsList.IndexOf(prop) + 1].Value;

                    if (ignoreObjectIfOnePropertyHasNoValue)
                    {
                        if (currentCellValue != null && currentCellValue.Trim() != "")
                        {
                            obj.GetType().GetProperty(prop.Name)?.SetValue(obj, Convert.ChangeType(currentCellValue, prop.PropertyType));
                        }
                    }
                    else
                    {
                        obj.GetType().GetProperty(prop.Name)?.SetValue(obj, Convert.ChangeType(currentCellValue, prop.PropertyType));
                    }
                }

                if ( ignoreObjectIfOnePropertyHasNoValue )
                { 
                    if ( !IsOneOfObjectPropertiesNull(obj) )
                    { 
                        result.Add(obj);
                    }
                }
                else
                {
                    result.Add(obj);
                }

            }

            return result;
        }
        private static List<T> Select<T>(WorkSheet workSheet, Expression<Func<T, object>> usedPropertiesExpression , int headerRowIndex, bool ignoreObjectIfOnePropertyHasNoValue) where  T : class, new()
        {
            var usedProperties            = GetUsedProperties( usedPropertiesExpression );
            var headerTitlesWithIndexes   = GetHeaderTitlesWithIndexes( workSheet , headerRowIndex);
            var cells                     = workSheet.Content.Cells;
            var lastUsedRowInTheWorkSheet = workSheet.Content.LastRow;
            var result                    = new List<T>();

            for (var i = headerRowIndex + 2; i <= lastUsedRowInTheWorkSheet; i++)
            {
                var row = cells[i];
                var obj = new T();

                foreach (var prop in usedProperties)
                {
                    var currentPropertyColumnIndex = headerTitlesWithIndexes[prop.Name];
                    var currentCellValue           = row[i , currentPropertyColumnIndex + 1].Value;

                    if (ignoreObjectIfOnePropertyHasNoValue)
                    {
                        if (currentCellValue != null && currentCellValue.Trim() != "")
                        {
                            obj.GetType().GetProperty(prop.Name)?.SetValue(obj, Convert.ChangeType(currentCellValue, prop.PropertyType));
                        }
                    }
                    else
                    {
                        obj.GetType().GetProperty(prop.Name)?.SetValue(obj, Convert.ChangeType(currentCellValue, prop.PropertyType));
                    }
                }

                if (ignoreObjectIfOnePropertyHasNoValue)
                {
                    if (!IsOneOfObjectPropertiesNull(obj ,usedPropertiesExpression))
                    {
                        result.Add(obj);
                    }
                }
                else
                {
                    result.Add(obj);
                }
            }

            return result;
        }
        private static List<T> Select<T, TSheetColumnsMap>(WorkSheet workSheet, int headerRowIndex, bool ignoreObjectIfOnePropertyHasNoValue) where  T : class, new() where  TSheetColumnsMap : ISheetColumnsMap<T>, new()
        {
            var propertiesOfT             = typeof(T).GetProperties();
            var cells                     = workSheet.Content.Cells;
            var lastUsedRowInTheWorkSheet = workSheet.Content.LastRow;
            var result                    = new List<T>();
            var columnsMap                = new TSheetColumnsMap().CreateMap();
            var headersIndexes            = GetSheetHeadersIndexes<T>(worksheet: workSheet, headersRowIndex: 0);

            for (var i = headerRowIndex + 2; i <= lastUsedRowInTheWorkSheet; i++)
            {
                var row = cells[i];
                var obj = new T();

                foreach (var prop in propertiesOfT)
                {
                    if (columnsMap.ContainsKey(prop.Name))
                    {
                        var propHeader = columnsMap[prop.Name];
                        var propHeaderIndex = headersIndexes[propHeader];
                        var currentCellValue = row[i, propHeaderIndex + 1].Value;

                        if (ignoreObjectIfOnePropertyHasNoValue)
                        {
                            if (currentCellValue != null && currentCellValue.Trim() != "")
                            {
                                obj.GetType().GetProperty(prop.Name)?.SetValue(obj, Convert.ChangeType(currentCellValue, prop.PropertyType));
                            }
                        }
                        else
                        {
                            obj.GetType().GetProperty(prop.Name)?.SetValue(obj, Convert.ChangeType(currentCellValue, prop.PropertyType));
                        }
                    }
                }

                if (ignoreObjectIfOnePropertyHasNoValue)
                {
                    if (!IsOneOfObjectPropertiesNull(obj))
                    {
                        result.Add(obj);
                    }
                }
                else
                {
                    result.Add(obj);
                }
            }

            return result;
        }
        private static List<T> Select<T, TSheetColumnsMap>(WorkSheet workSheet, Expression<Func<T, object>> usedPropertiesExpression , int headerRowIndex, bool ignoreObjectIfOnePropertyHasNoValue) where  T : class, new() where  TSheetColumnsMap : ISheetColumnsMap<T>, new()
        {
            var usedProperties            = GetUsedProperties( usedPropertiesExpression );
            var columnsMap                = new TSheetColumnsMap().CreateMap();
            var headerTitlesWithIndexes   = GetHeaderTitlesWithIndexes( workSheet , headerRowIndex);
            var cells                     = workSheet.Content.Cells;
            var lastUsedRowInTheWorkSheet = workSheet.Content.LastRow;
            var result                    = new List<T>();

            for (var i = headerRowIndex + 2; i <= lastUsedRowInTheWorkSheet; i++)
            {
                var row = cells[i];
                var obj = new T();

                foreach (var prop in usedProperties)
                {
                    if ( columnsMap.ContainsKey( prop.Name ) )
                    {
                        var currentPropertyHeader            = columnsMap[prop.Name];
                        var currentPropertyHeaderColumnIndex = headerTitlesWithIndexes[currentPropertyHeader];
                        var currentCellValue                 = row[i , currentPropertyHeaderColumnIndex + 1].Value;

                        if ( ignoreObjectIfOnePropertyHasNoValue )
                        {
                            
                            if ( currentCellValue != null && currentCellValue.Trim() != "" )
                            { 
                                obj.GetType().GetProperty( prop.Name )?.SetValue( obj , Convert.ChangeType( currentCellValue , prop.PropertyType ) );
                            }
                           
                        }
                        else
                        { 
                            obj.GetType().GetProperty( prop.Name )?.SetValue( obj , Convert.ChangeType( currentCellValue , prop.PropertyType ) );
                        }
                    }
                }

                if (ignoreObjectIfOnePropertyHasNoValue)
                {
                    if (!IsOneOfObjectPropertiesNull(obj, usedPropertiesExpression))
                    {
                        result.Add(obj);
                    }
                }
                else
                {
                    result.Add(obj);
                }
            }

            return result;
        }

        private static void _clearRange( WorkSheet workSheet , string range )
        {
            workSheet.Content.Range[range].ClearContents();
        }

        private static void _StylingRange(WorkSheet workSheet, string range, RangeStyle style )
        {
            workSheet.Content.Range[range].Style.HorizontalAlignment = style.HorizontalAlign;
            workSheet.Content.Range[range].Style.VerticalAlignment   = style.VerticalAlign;

            workSheet.Content.Range[range].Style.Font.Color          = ColorTranslator.FromHtml(style.FontColor);
            workSheet.Content.Range[range].Style.Interior.Color      = ColorTranslator.FromHtml(style.BackColor);

            workSheet.Content.Range[range].Style.Font.Size           = style.FontSize;
            workSheet.Content.Range[range].Style.Font.IsBold         = style.IsFontBold;
            workSheet.Content.Range[range].Style.Font.IsItalic       = style.IsFontItalic;
            workSheet.Content.Range[range].Style.Font.Underline      = style.FontUnderline;

            workSheet.Content.Range[range].Style.Borders.LineStyle   = style.BordersLineStyle;
            workSheet.Content.Range[range].Style.Borders.Color       = ColorTranslator.FromHtml( style.BordersColor );

            workSheet.Content.Range[range].RowHeight                 = style.RowHeight;
        }
        private static void _StylingRange(WorkSheet workSheet, int startRow, int startColumn, int lastRow, int lastColumn, RangeStyle style)
        {
            workSheet.Content.Range[startRow, startColumn, lastRow, lastColumn].Style.HorizontalAlignment = style.HorizontalAlign;
            workSheet.Content.Range[startRow, startColumn, lastRow, lastColumn].Style.VerticalAlignment   = style.VerticalAlign;

            workSheet.Content.Range[startRow , startColumn , lastRow , lastColumn].Style.Font.Color = ColorTranslator.FromHtml( style.FontColor );
            workSheet.Content.Range[startRow, startColumn, lastRow, lastColumn].Style.Interior.Color    = ColorTranslator.FromHtml(style.BackColor);

            workSheet.Content.Range[startRow, startColumn, lastRow, lastColumn].Style.Font.Size         = style.FontSize;
            workSheet.Content.Range[startRow, startColumn, lastRow, lastColumn].Style.Font.IsBold       = style.IsFontBold;
            workSheet.Content.Range[startRow, startColumn, lastRow, lastColumn].Style.Font.IsItalic     = style.IsFontItalic;
            workSheet.Content.Range[startRow, startColumn, lastRow, lastColumn].Style.Font.Underline    = style.FontUnderline;

            workSheet.Content.Range[startRow, startColumn, lastRow, lastColumn].Style.Borders.LineStyle = style.BordersLineStyle;
            workSheet.Content.Range[startRow, startColumn, lastRow, lastColumn].Style.Borders.Color     = ColorTranslator.FromHtml(style.BordersColor);

            workSheet.Content.Range[startRow, startColumn, lastRow, lastColumn].RowHeight               = style.RowHeight;
        }


        private static void _MergeRange( WorkSheet workSheet , string range, string value, HorizontalAlignType horizontalAlignType, VerticalAlignType verticalAlignType, string fontColor, string backColor)
        {
            workSheet.Content.Range[range].Merge( true );
            workSheet.Content.Range[range].Value                     = value;
            workSheet.Content.Range[range].Style.HorizontalAlignment = horizontalAlignType;
            workSheet.Content.Range[range].Style.VerticalAlignment   = verticalAlignType;
            workSheet.Content.Range[range].Style.Font.Color          = ColorTranslator.FromHtml( fontColor );
            workSheet.Content.Range[range].Style.Interior.Color      = ColorTranslator.FromHtml( backColor );
        } 
        private static void _MergeRange( WorkSheet workSheet , string range, string value, RangeStyle style)
        {
            workSheet.Content.Range[range].Merge( true );
            workSheet.Content.Range[range].Value = value;

            _StylingRange( workSheet , range , style );
        }
        private static void _MergeRangeAndSetValue( WorkSheet workSheet, string range, string value, HorizontalAlignType horizontalAlignType, VerticalAlignType verticalAlignType, string fontColor, string backColor  )
        {
            workSheet.Content.Range[range].Merge( true );
            workSheet.Content.Range[range].Value                     = value;
            workSheet.Content.Range[range].Style.HorizontalAlignment = horizontalAlignType;
            workSheet.Content.Range[range].Style.VerticalAlignment   = verticalAlignType;
            workSheet.Content.Range[range].Style.Font.Color          = ColorTranslator.FromHtml( fontColor );
            workSheet.Content.Range[range].Style.Interior.Color      = ColorTranslator.FromHtml( backColor );
        } 
        private static void _MergeRangeAndSetValue( WorkSheet workSheet, string range, string value, RangeStyle style)
        {
            workSheet.Content.Range[range].Merge( true );
            workSheet.Content.Range[range].Value = value;

            _StylingRange( workSheet , range , style );
        }

        private static void _DataHeaderStyle(WorkSheet workSheet, RangeStyle style, int headerRowIndex)
        {
            var firstColumnIndex = workSheet.Content.FirstColumn;
            var lastColumnIndex  = workSheet.Content.LastColumn;

            _StylingRange( workSheet , headerRowIndex + 1 , firstColumnIndex , headerRowIndex + 1 , lastColumnIndex , style );
        }
        private static void _DataBodyStyle(WorkSheet workSheet, RangeStyle style, int bodyStartRowIndex)
        {
            var firstColumnIndex = workSheet.Content.FirstColumn;
            var lastColumnIndex  = workSheet.Content.LastColumn;
            var lastRowIndex     = workSheet.Content.LastRow;

            // Body styling
            _StylingRange( workSheet , bodyStartRowIndex + 1 , firstColumnIndex , lastRowIndex , lastColumnIndex , style );
        }

        private static void _DeleteLastUsedRow( WorkSheet workSheet )
        {
            var lastUsedRowIndex = workSheet.Content.LastRow;
            workSheet.Content.DeleteRow( lastUsedRowIndex );
        }

        #endregion


        #region Public methods

        /// <summary>
        /// Fill in the worksheet with data
        /// </summary>
        /// <typeparam name="T">Type of data</typeparam>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="data">Data to be stored</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet SetData<T>(this WorkSheet workSheet, List<T> data)
        {
            ChangeTypeOfWorkSheet<T>(ref workSheet, data);
            PrepareTheWorkSheetHeaders(ref workSheet);
            PrepareTheWorkSheetData(ref workSheet, data);
            StylingTheWorkSheet(ref workSheet, data.Count);
            return workSheet;
        }

        /// <summary>
        /// Fill in the worksheet with data
        /// </summary>
        /// <typeparam name="T">Type of data</typeparam>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="data">Data to be stored</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet SetData<T, TSheetColumnsMap>(this WorkSheet workSheet, List<T> data) where T : class where TSheetColumnsMap : ISheetColumnsMap<T>, new()
        {
            ChangeTypeOfWorkSheet<T>(ref workSheet, data);
            PrepareTheWorkSheetHeaders<T, TSheetColumnsMap>(ref workSheet);
            PrepareTheWorkSheetData(ref workSheet, data);
            StylingTheWorkSheet(ref workSheet, data.Count);
            return workSheet;
        }

        /// <summary>
        /// Fill in the worksheet with data
        /// </summary>
        /// <typeparam name="T">Type of data</typeparam>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="data">Data to be stored</param>
        /// <param name="columnHeaders">Custom column headers text/titles</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet SetData<T>(this WorkSheet workSheet, List<T> data, List<string> columnHeaders)
        {
            ChangeTypeOfWorkSheet<T>(ref workSheet, data);
            PrepareTheWorkSheetHeaders(ref workSheet, columnHeaders);
            PrepareTheWorkSheetData(ref workSheet, data);
            StylingTheWorkSheet(ref workSheet, data.Count);
            return workSheet;
        }


        /// <summary>
        /// Protect the worksheet with a password
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="password">Password</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet Protect(this WorkSheet workSheet, string password)
        {
            workSheet.Content.Protect(password);
            return workSheet;
        }

        /// <summary>
        /// Protect the worksheet with a password
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="password">Password</param>
        /// <param name="protectionType">Represent worksheet protection flags</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet Protect(this WorkSheet workSheet, string password, SheetProtectionType protectionType)
        {
            workSheet.Content.Protect(password, protectionType);
            return workSheet;
        }


        /// <summary>
        /// Remove worksheet protection
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet Unprotect(this WorkSheet workSheet)
        {
            workSheet.Content.Unprotect();
            return workSheet;
        }

        /// <summary>
        /// Remove worksheet protection
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="password">Current worksheet password</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet Unprotect(this WorkSheet workSheet, string password)
        {
            workSheet.Content.Unprotect(password);
            return workSheet;
        }

        /// <summary>
        /// Set worksheet tab color
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="colorAsHex">Color as Hexadecimal</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet TabColor(this WorkSheet workSheet, string colorAsHex)
        {
            workSheet.Content.TabColor = ColorTranslator.FromHtml(colorAsHex);
            return workSheet;
        }

        /// <summary>
        /// Set worksheet tab color
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="color">Color as <see cref="Color"/> provided colors</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet TabColor(this WorkSheet workSheet, Color color)
        {
            workSheet.Content.TabColor = color;
            return workSheet;
        }


        /// <summary>
        /// Freeze worksheet table header panes
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet FreezeHeadersPane(this WorkSheet workSheet)
        {
            workSheet.Content.FreezePanes(2, 1);
            return workSheet;
        }

        /// <summary>
        /// Freeze worksheet table header panes
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="rowIndex">Worksheet data table header row index</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet FreezeHeadersPane(this WorkSheet workSheet, int rowIndex)
        {
            workSheet.Content.FreezePanes(rowIndex + 2, 1);
            return workSheet;
        }


        /// <summary>
        /// Delete a specific column from the worksheet
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="columnIndex">Index of column to be deleted</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet DeleteColumn(this WorkSheet workSheet, int columnIndex)
        {
            workSheet.Content.DeleteColumn(columnIndex);
            return workSheet;
        }

        /// <summary>
        /// Delete a specific column from the worksheet
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="columnIndex">Index of column to be deleted</param>
        /// <param name="count">Number of columns to be deleted, starting from the specified column</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet DeleteColumn(this WorkSheet workSheet, int columnIndex, int count)
        {
            workSheet.Content.DeleteColumn(columnIndex, count);
            return workSheet;
        }


        /// <summary>
        /// Delete a specific row from the worksheet
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="rowIndex">Index of row to be deleted</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet DeleteRow(this WorkSheet workSheet, int rowIndex)
        {
            workSheet.Content.DeleteRow(rowIndex + 1);
            return workSheet;
        }

        /// <summary>
        /// Delete a specific row from the worksheet
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="rowIndex">Index of row to be deleted</param>
        /// <param name="count">Number of rows to be deleted, starting from the specified row</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet DeleteRow(this WorkSheet workSheet, int rowIndex, int count)
        {
            workSheet.Content.DeleteRow(rowIndex + 1, count);
            return workSheet;
        }


        /// <summary>
        /// Asynchronously select data from the worksheet
        /// </summary>
        /// <typeparam name="T">Type of data</typeparam>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="headerRowIndex">Index of the data table header row</param>
        /// <param name="ignoreObjectIfOnePropertyHasNoValue">Determine if want to ignore objects that one of its properties has no value</param>
        /// <returns><see cref="Task{TResult}"/></returns>
        public static Task<List<T>> SelectAsync<T>(this WorkSheet workSheet, int headerRowIndex = 0, bool ignoreObjectIfOnePropertyHasNoValue = false) where T : class, new()
        {
            return Task.Factory.StartNew(() => Select<T>(workSheet, headerRowIndex, ignoreObjectIfOnePropertyHasNoValue));
        }

        /// <summary>
        /// Asynchronously select data from the worksheet and determine just one or bunche properties to be selected
        /// </summary>
        /// <typeparam name="T">Type of data</typeparam>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="usedPropertiesExpression">Properties to be selected from the worksheet and set its values</param>
        /// <param name="headerRowIndex">Index of the data table header row</param>
        /// <param name="ignoreObjectIfOnePropertyHasNoValue">Determine if want to ignore objects that one of its properties has no value</param>
        /// <returns><see cref="Task{TResult}"/></returns>
        public static Task<List<T>> SelectAsync<T>(this WorkSheet workSheet, Expression<Func<T, object>> usedPropertiesExpression, int headerRowIndex = 0, bool ignoreObjectIfOnePropertyHasNoValue = false) where T : class, new()
        {
            return Task.Factory.StartNew(() => Select<T>(workSheet, usedPropertiesExpression, headerRowIndex, ignoreObjectIfOnePropertyHasNoValue));
        }

        /// <summary>
        /// Asynchronously select data from the worksheet
        /// </summary>
        /// <typeparam name="T">Type of data</typeparam>
        /// <typeparam name="TSheetColumnsMap">An implementation of <see cref="ISheetColumnsMap{T}"/></typeparam>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="headerRowIndex">Index of the data table header row</param>
        /// <param name="ignoreObjectIfOnePropertyHasNoValue">Determine if want to ignore objects that one of its properties has no value</param>
        /// <returns><see cref="Task{TResult}"/></returns>
        public static Task<List<T>> SelectAsync<T, TSheetColumnsMap>(this WorkSheet workSheet, int headerRowIndex = 0, bool ignoreObjectIfOnePropertyHasNoValue = false) where T : class, new() where TSheetColumnsMap : ISheetColumnsMap<T>, new()
        {
            return Task.Factory.StartNew(() => Select<T>(workSheet, headerRowIndex, ignoreObjectIfOnePropertyHasNoValue));
        }

        /// <summary>
        /// Asynchronously select data from the worksheet and determine just one or bunche properties to be selected
        /// </summary>
        /// <typeparam name="T">Type of data</typeparam>
        /// <typeparam name="TSheetColumnsMap">An implementation of <see cref="ISheetColumnsMap{T}"/></typeparam>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="usedPropertiesExpression">Properties to be selected from the worksheet and set its values</param>
        /// <param name="headerRowIndex">Index of the data table header row</param>
        /// <param name="ignoreObjectIfOnePropertyHasNoValue">Determine if want to ignore objects that one of its properties has no value</param>
        /// <returns><see cref="Task{TResult}"/></returns>
        public static Task<List<T>> SelectAsync<T, TSheetColumnsMap>(this WorkSheet workSheet, Expression<Func<T, object>> usedPropertiesExpression, int headerRowIndex = 0, bool ignoreObjectIfOnePropertyHasNoValue = false) where T : class, new() where TSheetColumnsMap : ISheetColumnsMap<T>, new()
        {
            return Task.Factory.StartNew(() => Select<T, TSheetColumnsMap>(workSheet, usedPropertiesExpression, headerRowIndex, ignoreObjectIfOnePropertyHasNoValue));
        }


        /// <summary>
        /// Lock a cell or a range of cells
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="range">Cell or range of cells name</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet LockRange(this WorkSheet workSheet, string range)
        {
            workSheet.Content.Range[range].Style.Locked = true;
            return workSheet;
        }

        /// <summary>
        /// Unlock a cell or a range of cells
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="range">Cell or range of cells name</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet UnlockRange(this WorkSheet workSheet, string range)
        {
            workSheet.Content.Range[range].Style.Locked = false;
            return workSheet;
        }


        /// <summary>
        /// Add a range of cells that allow editing when worksheet protected
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="range">Cell or range of cells name</param>
        /// <param name="title">Title</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet AllowEditRange(this WorkSheet workSheet, string range, string title = "allowed range")
        {
            workSheet.Content.AddAllowEditRange(title, workSheet.Content.Range[range]);
            return workSheet;
        }


        /// <summary>
        /// Asynchronously save the worksheet as a PDF file
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="pdfPath">Path to save the PDF in</param>
        public static Task ToPdfAsync(this WorkSheet workSheet, string pdfPath)
        {
            return Task.Factory.StartNew(() => workSheet.Content.SaveToPdf(pdfPath, FileFormat.PDF));
        }

        /// <summary>
        /// Save the worksheet as a Image file
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="pdfPath">Path to save the image in</param>
        public static Task ToImage(this WorkSheet workSheet, string pdfPath)
        {
            return Task.Factory.StartNew(() => workSheet.Content.SaveToImage(pdfPath, ImageFormat.Jpeg));
        }

        /// <summary>
        /// Save the worksheet as a Image file
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="pdfPath">Path to save the image in</param>
        /// <param name="format">Image format</param>
        public static Task ToImage(this WorkSheet workSheet, string pdfPath, ImageFormat format)
        {
            return Task.Factory.StartNew(() => workSheet.Content.SaveToImage(pdfPath, format));
        }

        /// <summary>
        /// Asynchronously save the worksheet as an HTML file
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="path">Path to save the .html file in</param>
        /// <returns><see cref="Task"/></returns>
        public static Task ToHtmlAsync(this WorkSheet workSheet, string path) => Task.Factory.StartNew(() => workSheet.Content.SaveToHtml(path));


        /// <summary>
        /// Groups columns
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="firstColumn">The first column index to be grouped</param>
        /// <param name="lastColumn">The last column index to be grouped</param>
        /// <param name="isCollapsed">Indicates whether group should be collapsed</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet GroupColumns(this WorkSheet workSheet, int firstColumn, int lastColumn, bool isCollapsed = true)
        {
            workSheet.Content.GroupByColumns(firstColumn + 1, lastColumn + 1, isCollapsed);
            return workSheet;
        }

        /// <summary>
        /// Ungroups columns
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="firstColumn">The first column index to be ungrouped</param>
        /// <param name="lastColumn">The last column index to be ungrouped</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet UngroupColumns(this WorkSheet workSheet, int firstColumn, int lastColumn)
        {
            workSheet.Content.UngroupByColumns(firstColumn + 1, lastColumn + 1);
            return workSheet;
        }

        /// <summary>
        /// Groups rows
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="firstRow">The first row index to be grouped</param>
        /// <param name="lastRow">The last row index to be grouped</param>
        /// <param name="isCollapsed">Indicates whether group should be collapsed</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet GroupRows(this WorkSheet workSheet, int firstRow, int lastRow, bool isCollapsed = true)
        {
            workSheet.Content.GroupByRows(firstRow + 1, lastRow + 1, isCollapsed);
            return workSheet;
        }

        /// <summary>
        /// Ungroups rows
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="firstRow">The first row index to be ungrouped</param>
        /// <param name="lastRow">The last row index to be ungrouped</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet UngroupRows(this WorkSheet workSheet, int firstRow, int lastRow)
        {
            workSheet.Content.UngroupByRows(firstRow + 1, lastRow + 1);
            return workSheet;
        }


        /// <summary>
        /// Set the data table column header with a custom colors
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="backColor">The background color in hexadecimal</param>
        /// <param name="fontColor">The font color in hexadecimal</param>
        /// <param name="headerRowIndex">The index of header row</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet DataHeaderColors(this WorkSheet workSheet, string backColor, string fontColor, int headerRowIndex = 0)
        {
            StylingTheHeader(ref workSheet, backColor, fontColor, headerRowIndex);
            return workSheet;
        }

        /// <summary>
        /// Styling worksheet data table header cells
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="style">Range style</param>
        /// <param name="headerRowIndex">The index of header row</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet DataHeaderStyle( this WorkSheet workSheet , RangeStyle style , int headerRowIndex = 0 )
        {
            _DataHeaderStyle( workSheet , style , headerRowIndex );
            return workSheet;
        }


        /// <summary>
        /// Set the data table body with a custom colors
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="style">Range style</param>
        /// <param name="bodyStartRowIndex">The index of row that body start from</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet DataBodyStyle(this WorkSheet workSheet, RangeStyle style, int bodyStartRowIndex = 1)
        {
            _DataBodyStyle( workSheet , style , bodyStartRowIndex );
            return workSheet;
        } 
        
        /// <summary>
        /// Set the data table body with a custom colors
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="backColor">The background color in hexadecimal</param>
        /// <param name="fontColor">The font color in hexadecimal</param>
        /// <param name="bodyStartRowIndex">The index of row that body start from</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet DataBodyColors(this WorkSheet workSheet, string backColor, string fontColor, int bodyStartRowIndex = 1)
        {
            StylingTheBody(ref workSheet, backColor, fontColor, bodyStartRowIndex);
            return workSheet;
        }


        /// <summary>
        /// Clear a range of cells contents
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="range">Cell or range of cells name</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet ClearRange(this WorkSheet workSheet, string range)
        {
            _clearRange(workSheet, range);
            return workSheet;
        }


        /// <summary>
        /// Creates a merged cell from a specified range
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="range">Cell or range of cells name</param>
        /// <param name="value">Value of merged range</param>
        /// <param name="horizontalAlignType">Text horizontal alignment</param>
        /// <param name="verticalAlignType">Text vertical alignment</param>
        /// <param name="fontColor">Font color</param>
        /// <param name="backColor">Background color</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet MergeRange(this WorkSheet workSheet , string range, string value="", HorizontalAlignType horizontalAlignType = HorizontalAlignType.Left, VerticalAlignType verticalAlignType = VerticalAlignType.Bottom, string fontColor = "#000000", string backColor = "#FFFFFF")
        {
            _MergeRange( workSheet , range , value , horizontalAlignType , verticalAlignType , fontColor , backColor );
            return workSheet;
        }

        /// <summary>
        /// Creates a merged cell from a specified range
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="range">Cell or range of cells name</param>
        /// <param name="style">Range style</param>
        /// <param name="value">Value of merged range</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet MergeRange(this WorkSheet workSheet , string range, RangeStyle style, string value= "")
        {
            _MergeRange( workSheet , range , value , style );
            return workSheet;
        }

      
        /// <summary>
        /// Creates a merged cell from a specified range and set its value
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="range">Cell or range of cells name</param>
        /// <param name="value">Value</param>
        /// <param name="horizontalAlignType">Text horizontal alignment</param>
        /// <param name="verticalAlignType">Text vertical alignment</param>
        /// <param name="fontColor">Font color</param>
        /// <param name="backColor">Background color</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet MergeRangeAndSetValue(this WorkSheet workSheet, string range, string value, HorizontalAlignType horizontalAlignType = HorizontalAlignType.Left, VerticalAlignType verticalAlignType = VerticalAlignType.Bottom, string fontColor = "#000000", string backColor = "#FFFFFF")
        {
            _MergeRangeAndSetValue( workSheet , range , value , horizontalAlignType , verticalAlignType , fontColor , backColor );
            return workSheet;
        }

        /// <summary>
        /// Creates a merged cell from a specified range and set its value
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="range">Cell or range of cells name</param>
        /// <param name="value">Value</param>
        /// <param name="style">Range style</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet MergeRangeAndSetValue(this WorkSheet workSheet, string range, string value, RangeStyle style)
        {
            _MergeRangeAndSetValue( workSheet , range , value , style );
            return workSheet;
        }


        /// <summary>
        /// Styling a range of cells
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <param name="range">Cell or range of cells name</param>
        /// <param name="style">Range style</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet StylingRange(this WorkSheet workSheet, string range, RangeStyle style)
        {
            _StylingRange( workSheet , range , style );
            return workSheet;
        }

        /// <summary>
        /// Delete the last used row
        /// </summary>
        /// <param name="workSheet">Represent <see cref="WorkSheet"/> object</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public static WorkSheet DeleteLastUsedRow( this WorkSheet workSheet )
        {
            _DeleteLastUsedRow( workSheet );
            return workSheet;
        }

        #endregion

    }
}
