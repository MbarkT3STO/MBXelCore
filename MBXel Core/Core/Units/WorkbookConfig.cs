using MBXel_Core.Core.Abstraction;
using MBXel_Core.Enums;

using Spire.Xls;

namespace MBXel_Core.Core.Units
{
    public class WorkbookConfig : IWorkbookConfig
    {
        public string       Path        { get; set; }
        public int          SheetsCount { get; set; }
        public XLExtension  Extension   { get; set; }
        public ExcelVersion Version     { get; set; }
        public string       Password    { get; set; }
    }
}
