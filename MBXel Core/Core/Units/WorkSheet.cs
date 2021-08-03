using MBXel_Core.Core.Abstraction;

using Spire.Xls;

using System;

namespace MBXel_Core.Core.Units
{
    public class WorkSheet : IWorkSheet
    {
        public Type Type { get; set; }
        public Worksheet Content { get; set; }
        public string Name { get; set; }

        public void SetName(string sheetName)
        {
            Name = sheetName;
            Content.Name = Name;
        } 
    }
}
