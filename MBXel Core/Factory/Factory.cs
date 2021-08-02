using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;
using MBXel_Core.Core.Abstraction;
using MBXel_Core.Core.Units;

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
