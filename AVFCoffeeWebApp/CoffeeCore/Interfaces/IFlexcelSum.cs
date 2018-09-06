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
            Double productionQuintales, Double transportCostSoles, Double costPriceSolesPerQuintal, Double expSolesChem, Double expSolesOrg);

        void SaveUserInputs(string id, ChartInputDTO chartInputDTO);

        LoginInfoDTO GetUserInputs(String number);

        ChartDataDTO SaveUserOutputs(string id, ChartDataDTO chartDataDTO);

        UserInfoDTO SaveUserInfo(UserInfoDTO userInfoDTO);

        UserInfoDTO UpdateUserInfo(UserInfoDTO userInfoDTO);
    }
}
