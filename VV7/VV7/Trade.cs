using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VV7
{
    class Trade
    {
        public string TradeType { get; set;  }
        public string Symbol { get; set; }
        public string TradeDate { get; set; }
        public int Quantity { get; set; }
        public double Price { get; set; }
        public double Commision { get; set; }
    }
}
