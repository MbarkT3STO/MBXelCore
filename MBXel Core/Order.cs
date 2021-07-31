using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MBXel_Core
{
    class Order
    {
        public int ID { get; set; }
        public string Client { get; set; }
        public string Product { get; set; }
        public int Total { get; set; }


        public Order(int _ID, string _Client, string _Product, int _Total)
        {
            ID = _ID;
            Client = _Client;
            Product = _Product;
            Total = _Total;
        }

        public Order()
        {

        }
    }
}
