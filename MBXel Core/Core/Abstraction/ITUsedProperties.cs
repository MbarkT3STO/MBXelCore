using System;
using System.Collections.Generic;

namespace MBXel_Core.Core.Abstraction
{
    public interface ITUsedProperties<in T> where  T : class
    {
        IEnumerable<string> GetUsedProperties();
    }
}