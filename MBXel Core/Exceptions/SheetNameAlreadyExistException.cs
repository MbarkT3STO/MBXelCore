using System;

namespace MBXel_Core.Exceptions
{
    public class SheetNameAlreadyExistException : Exception
    {
        public SheetNameAlreadyExistException(string message= "A worksheet with the same name already exist") : base(message)
        {

        }
    }
}
