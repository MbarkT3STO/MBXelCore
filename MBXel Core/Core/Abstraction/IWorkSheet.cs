using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Spire.Xls;

namespace MBXel_Core.Core.Abstraction
{
    public interface IWorkSheet
    {
        Type Type { get; set; }
        Worksheet Content { get; set; }
        string Name { get; set; }

        void SetName(string sheetName);
    }
}
