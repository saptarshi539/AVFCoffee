using CoffeeInfrastructure.Flexcel;
using System;
using System.Collections.Generic;
using System.Text;
using CoffeeCore.DTO;

namespace CoffeeCore.Interfaces
{
    public interface IFlexcelsum
    {

        ChartDataDTO getOutputFromExcel(Double earlyHectares, Double peakHectares, Double oldHectares, bool conventional, bool organic, bool transition, Double workerSalarySoles,
            Double productionQuintales, Double transportCostSoles, Double costPriceSolesPerQuintal);

        void SaveUserInputs(string id, ChartInputDTO chartInputDTO);

        ChartInputDTO GetUserInputs(String id);
    }
}
