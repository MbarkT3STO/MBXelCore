using MBXel_Core.Core.Abstraction;

using System.Collections.Generic;

namespace MBXel_Core
{
    class ImplementationOfISheetColumnsMapExample : ISheetColumnsMap<Order>
    {
        public Dictionary<string , string> CreateMap()
        {
            var result = new Dictionary<string , string>();

            result.Add(nameof(Order.ID), "ID");
            result.Add(nameof(Order.Client), "Client Name");
            result.Add(nameof(Order.Product), "Product");
            result.Add(nameof(Order.Total), "Amount");

            return result;
        }
    }
}
