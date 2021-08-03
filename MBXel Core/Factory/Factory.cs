using MBXel_Core.Core.Units;

using Spire.Xls;

using System.Collections.Generic;

namespace MBXel_Core.Factory
{
    public class Factory
    {
        /// <summary>
        /// Create a new <see cref="Workbook"/> object
        /// </summary>
        /// <param name="numberOfSheets">Number of sheets should be in the workbook</param>
        /// <returns><see cref="Workbook"/></returns>
        public Spire.Xls.Workbook CreateWorkbook(int numberOfSheets)
        {
            var workBook = new Spire.Xls.Workbook();
            workBook.CreateEmptySheets(numberOfSheets);
            return workBook;
        } 
        
        /// <summary>
        /// Create a new <see cref="Workbook"/> object
        /// </summary>
        /// <returns><see cref="Workbook"/></returns>
        public Spire.Xls.Workbook CreateWorkbook()
        {
            var workBook = new Spire.Xls.Workbook();
            return workBook;
        }

        /// <summary>
        /// Create a new worksheet
        /// </summary>
        /// <param name="workBook">Workbook to insert into</param>
        /// <param name="sheetName">Worksheet name</param>
        /// <returns><see cref="WorkSheet"/></returns>
        public void CreateWorkSheet(ref Spire.Xls.Workbook workBook, string sheetName)
        {

            if (sheetName != null)
                workBook.CreateEmptySheet(sheetName);
            else
                workBook.CreateEmptySheet();

        }

        /// <summary>
        /// Create a customized number of sheets
        /// </summary>
        /// <param name="numberOfSheets">Number of sheets to be created</param>
        /// <returns><see cref="List{WorkSheet}"/></returns>
        public List<WorkSheet> CreateWorkSheets(int numberOfSheets)
        {
            var sheets = new List<WorkSheet>();
            for (int i=0; i<numberOfSheets; i++)
            {
                sheets.Add(new WorkSheet());
            }

            return sheets;
        }
    }
}
