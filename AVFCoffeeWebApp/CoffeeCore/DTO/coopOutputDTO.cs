﻿using System;
using System.Collections.Generic;
using System.Text;

namespace CoffeeCore.DTO
{
    public class coopOutputDTO
    {
        public string coopId { get; set; }
        public Double variableCostUSPound { get; set; }
        public Double fixedCostUSPound { get; set; }
        public Double totalCostAndDeprUSPound { get; set; }
        public Double totalCostUSPound { get; set; }
        public Double breakEvenCostUSPound { get; set; }
    }
}
