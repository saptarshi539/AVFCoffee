using System;
using System.Collections.Generic;
using System.Text;

namespace CoffeeCore.Interfaces
{
    public interface IFuctionOutputProvider
    {
        Double getOutputFromExcel(Double earlyHectares, Double peakHectares, Double oldHectares, bool conventional, bool organic, bool transition, Double workerSalarySoles,
            Double productionQuintales, Double transportCostSoles, Double costPriceSolesPerQuintal, Double expSolesChem, Double expSolesOrg);
    }
}
