using System;

namespace MBXel_Core.Exceptions
{
    class HeadersPropertiesNotEqualsToDataPropertiesException : Exception
    {
        public HeadersPropertiesNotEqualsToDataPropertiesException() : base("Headers properties do not equal to type of Data to be exported properties")
        {

        }
    }
}
