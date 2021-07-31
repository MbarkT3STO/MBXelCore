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

        #region Private methods

        private void _StylingTheWorkSheet(ref Worksheet WSheet, int ColumnsNumber, int RowsNumber)
        {
            //Columns styling
            WSheet.Range["A1:BB1"].Style.Font.Size = 14;
            WSheet.Range["A1:BB1"].Style.Font.IsBold = true;
            WSheet.Range["A1:BB1"].Style.Font.Color = Color.White;
            WSheet.Range["A1:BB1"].Style.Interior.Color = ColorTranslator.FromHtml("#54a0ff");
            WSheet.Range["A1:BB1"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            WSheet.Range["A1:BB1"].Style.VerticalAlignment = VerticalAlignType.Center;

            //Rows styling
            WSheet.Range[$"A2:BB{RowsNumber + 1}"].Style.Font.Size = 14;
            WSheet.Range[$"A2:BB{RowsNumber + 1}"].Style.Font.Color = Color.White;
            WSheet.Range[$"A2:BB{RowsNumber + 1}"].Style.Interior.Color = ColorTranslator.FromHtml("#2ed573");
            WSheet.Range[$"A2:BB{RowsNumber + 1}"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            WSheet.Range[$"A2:BB{RowsNumber + 1}"].Style.VerticalAlignment = VerticalAlignType.Center;

            //Other Columns styling
            WSheet.AllocatedRange.AutoFitRows();
            WSheet.AllocatedRange.AutoFitColumns();

            //Other Rows styling
            WSheet.SetRowHeight(1, 30);
        }


        private async void _Export<T>(List<T> data, string path, Enums.XLExtension extension, ExcelVersion version)
        {
            //Get data parameter type properties
            PropertyInfo[] data_Properties = typeof(T).GetProperties();
            var data_PropertiesCount = data_Properties.Length;

            try
            {
                //Prepaire the Workbook and Worksheet
                var Wbook = new Workbook();
                var Wsheet = (Worksheet)Wbook.Worksheets[0];

                //Put data into worksheet
                await Task.Run(() =>
                {
                    for (int i = 0; i < data_PropertiesCount; i++)
                    {
                        Wsheet.Range[1, i + 1].Text = data_Properties[i].Name;
                    }

                    int rowIndex = 2;

                    foreach (T d in data)
                    {
                        for (int i = 0; i < data_PropertiesCount; i++)
                        {
                            Wsheet.Range[rowIndex, i + 1].Text = data_Properties[i].GetValue(d).ToString();
                        }

                        rowIndex++;
                    }
                });

                //Styling the worksheet
                await Task.Run(() => _StylingTheWorkSheet(ref Wsheet, data_PropertiesCount, data.Count));

                //Save the workbook 
                Wbook.SaveToFile(path + (extension is Enums.XLExtension.Xlsx ? ".xlsx" : ".xls"), version);
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                //Kill all the opened untitled EXCEL processes
                Process[] processes = Process.GetProcessesByName("EXCEL");

                foreach (Process p in processes)
                {
                    if (p.MainWindowTitle.Length == 0)
                    {
                        p.Kill();
                    }
                }
            }
        }

        #endregion


        /// <summary>
        /// Export a <see cref="List{T}"/> of data to an excel file
        /// </summary>
        /// <param name="data">Data to be exported</param>
        /// <param name="path">Path to be save in</param>
        /// <param name="extension">Excel file extension</param>
        /// <param name="version">Excel file version</param>
        /// <returns><see cref="bool"/></returns>
        public bool Export<T>(List<T> data, string path, Enums.XLExtension extension = Enums.XLExtension.Xlsx, ExcelVersion version = ExcelVersion.Version2016)
        {
            _Export(data, path, extension, version);

            return true;
        }

        /// <summary>
        /// Asynchronously, <inheritdoc cref="Export{T}(List{T}, string, Enums.XLExtension, ExcelVersion)"/>
        /// </summary>
        /// <inheritdoc cref="Export{T}(List{T}, string, Enums.XLExtension, ExcelVersion)"/>
        /// <returns><see cref="Task{TResult}"/></returns>
        public Task<bool> ExportAsync<T>(List<T> data, string path, Enums.XLExtension extension = Enums.XLExtension.Xlsx, ExcelVersion version = ExcelVersion.Version2016)
        {
            return Task.Factory.StartNew(() =>
            {
                _Export(data, path, extension, version);
                return true;
            });
        }
    }
}
