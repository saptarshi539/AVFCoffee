using System;
using System.Collections.Generic;
using System.Text;

namespace CoffeeCore.DTO
{
    public class ChartInputDTO
    {
        Double earlyHectares { get; set; }
        Double peakHectares { get; set; }
        Double oldHectares { get; set; }
        bool conventional { get; set; }
        bool organic { get; set; }
        bool transition { get; set; }
        Double workerSalarySoles { get; set; }
        Double productionQuintales { get; set; }
        Double transportCostSoles { get; set; }
        Double costPriceSolesPerQuintal { get; set; }
    }
}
