using Spire.Xls;

namespace MBXel_Core.Core.Abstraction
{
    public interface IWorkbookConfig
    {
        string              Path        { get; set; }
        int                 SheetsCount { get; set; }
        Enums.XLExtension   Extension   { get; set; }
        public ExcelVersion Version     { get; set; }
        string              Password    { get; set; }
    }
}