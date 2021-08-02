using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MBXel_Core.Core.Abstraction;

using Spire.Xls;

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
