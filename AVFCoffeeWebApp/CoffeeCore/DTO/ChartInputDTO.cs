using System;
using System.Collections.Generic;
using System.Text;

namespace CoffeeCore.DTO
{
    public class ChartInputDTO
    {
        public Double earlyHectares { get; set; }
        public Double peakHectares { get; set; }
        public Double oldHectares { get; set; }
        public bool conventional { get; set; }
        public bool organic { get; set; }
        public bool transition { get; set; }
        public Double workerSalarySoles { get; set; }
        public Double productionQuintales { get; set; }
        public Double transportCostSoles { get; set; }
        public Double costPriceSolesPerQuintal { get; set; }
        public Double expSolesChem { get; set; }
        public Double expSolesOrg { get; set; }
    }
}
