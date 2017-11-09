using CoffeeInfrastructure.Flexcel;
using System;
using System.Collections.Generic;
using System.Text;

namespace CoffeeCore.Interfaces
{
    public interface IFlexcelsum
    {
        String sumcells();

        ChartDataDTO getOutputFromExcel(Double earlyHectares, Double peakHectares, Double oldHectares, bool conventional, bool organic, bool transition, Double workerSalarySoles,
            Double productionQuintales, Double transportCostSoles, Double costPriceSolesPerQuintal);
    }
}
