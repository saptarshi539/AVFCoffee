using System;
using System.Collections.Generic;
using System.Text;

namespace CoffeeCore.DTO
{
    public class ProducerOutputSpanishDTO
    {
        public string userID { get; set; }
        public Double variableCostsUSHect { get; set; }
        public Double variableCostsSolesHect { get; set; }
        public Double totalCostUSHect { get; set; }
        public Double totalCostSolesHect { get; set; }
        public Double breakEvenCostUSPound { get; set; }
    }
}
