using System;

namespace MBXel_Core.Exceptions
{
    public class SheetNameNotExistException : Exception
    {
        public SheetNameNotExistException(string message= "The worksheet name provided doesn't exist") : base(message)
        {

        }
    }

}
