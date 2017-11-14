using System;
using System.Collections.Generic;
using System.Text;

namespace CoffeeCore.DTO
{
    public class ProducerOutputEnglishDTO
    {
        public string userID { get; set; }
        public Double variableCostsUSPound { get; set; }
        public Double fixedCostsUSPound { get; set; }
        public Double totalCostAndDeprUSPound { get; set; }
        public Double totalCostsUSPound { get; set; } 
        public Double breakEvenCostUSPound { get; set; }
    }
}
