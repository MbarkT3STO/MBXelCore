using System;
using System.Collections.Generic;
using MBXel_Core.Core.Abstraction;

namespace MBXel_Core
{
    internal class ImplementationOfIUsedPropertiesExample : ITUsedProperties<Order>
    {
        public IEnumerable<string> GetUsedProperties()
        {
            var usedPropertiesForOrder = new List<string>()
                                         {
                                             nameof( Order.Client ) ,
                                             nameof( Order.Product ) ,
                                             nameof( Order.Total ) ,
                                         };

            return usedPropertiesForOrder;
        }
    }
}