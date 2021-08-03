using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MBXel_Core.Exceptions
{
    public class SheetIndexNotExistException : Exception
    {
        public SheetIndexNotExistException(string message = "The worksheet index provided not exist") : base(message)
        {

        }
    }
}
