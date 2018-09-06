using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;

using CoffeeCore.Interfaces;
using CoffeeInfrastructure.Flexcel;
using CoffeeCore.DTO;
using CoffeeInfrastructure.Helpers;

namespace Coffee.APIControllers
{
    [Route("api/[controller]")]
    public class CellSumController : Controller
    {
        IFlexcelsum flexcelsum;
        IFarmer farmer;
       
        public CellSumController(IFlexcelsum _flexcelsum, IFarmer _farmer)
        {
            flexcelsum = _flexcelsum;
            farmer = _farmer;
        }

        //[Route("sum/{cellId:long}")]
        //[HttpGet]
        //public IActionResult CalculateSum(long cellId)
        //{
        //    try
        //    {
        //        var l = cellId;
        //        String sContent = flexcelsum.sumcells();
        //        return Ok(sContent);
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e.InnerException);
        //        return StatusCode(500);
        //    }

        //}

        [Route("calculate")]
        [HttpGet]
        [Produces("application/json")]
        public IActionResult GetChartValues(Double earlyHectares = 1.03, Double peakHectares = 1.94, Double oldHectares = 1.97, bool conventional = true, bool organic = false, 
            bool transition = false, Double workerSalarySoles = 16.16, Double productionQuintales = 14, Double transportCostSoles = 235.22, Double costPriceSolesPerQuintal = 556.51, 
            Double expSolesChem = 379.80, Double expSolesOrg = 379.80)
        {
            try
            {
                ChartDataDTO sContent = flexcelsum.getOutputFromExcel(earlyHectares,peakHectares,oldHectares,conventional, organic, transition, workerSalarySoles, productionQuintales,
                    transportCostSoles, costPriceSolesPerQuintal, expSolesChem, expSolesOrg);
                return Ok(sContent);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return StatusCode(500);
            }

        }

        [Route("saveinput")]
        [HttpPost]
        [Produces("application/json; charset=utf-8")]
        public IActionResult PostInputs([FromBody]ChartInputDTO chartInputDTO)
        {
            try
            {
                //if (User.Identity.IsAuthenticated)
                //{
                    var id = chartInputDTO.phoneNumber; //"e661c05f-dc88-48c3-8026-3718143c56d8";//
                    flexcelsum.SaveUserInputs(id, chartInputDTO);
                //}
                //ChartDataDTO sContent = null;
                return Ok();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return StatusCode(500);
            }

        }

        [Route("saveoutput")]
        [HttpPost]
        [Produces("application/json; charset=utf-8")]
        public IActionResult PostOutputs([FromBody] ChartDataDTO chartDataDTO)
        {
            try
            {
                //if (User.Identity.IsAuthenticated)
                //{
                    var id = chartDataDTO.phoneNumber; //"e661c05f-dc88-48c3-8026-3718143c56d8";//
                    flexcelsum.SaveUserOutputs(id, chartDataDTO);
                //}
                //ChartDataDTO sContent = null;
                return Ok();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return StatusCode(500);
            }

        }

        //[Route("saveuser")]
        //[HttpPost]
        //[Produces("application/json")]
        //public IActionResult PostUserData([FromBody] UserInfoDTO userInfoDTO)
        //{
        //    try
        //    {
        //        if (User.Identity.IsAuthenticated)
        //        {
        //            var id = User.GetId(); //"e661c05f-dc88-48c3-8026-3718143c56d8";//
        //            flexcelsum.SaveUserInfo(id, userInfoDTO);
        //        }
        //        //ChartDataDTO sContent = null;
        //        return Ok();
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e.InnerException);
        //        return StatusCode(500);
        //    }

        //}
        [Route("FarmerLogin")]
        [HttpGet]
        [Produces("application/json")]
        public IActionResult AuthenticateFarmerLogin(String phoneNumber)
        {
            try
            {
                var status = false;
                if (phoneNumber != null)
                {
                    status = farmer.checkPhoneNumber(phoneNumber);
                }
                return Ok(status);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return StatusCode(500);
            }

        }


        [Route("getinput")]
        [HttpGet]
        [Produces("application/json")]
        public IActionResult GetInputs(String phoneNumber)
        {
            try
            {
                LoginInfoDTO output = new LoginInfoDTO();
                //if (User.Identity.IsAuthenticated)
                //{
                //var id = User.GetId();
                output = flexcelsum.GetUserInputs(phoneNumber);
                //}
                //ChartDataDTO sContent = null;
                return Ok(output);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return StatusCode(500);
            }

        }


        public LoginInfoDTO GetOutputStatus(string userid)
        {
            try
            {
                LoginInfoDTO output = new LoginInfoDTO();
                output = flexcelsum.GetUserInputs(userid);
                return output;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return null;
                //return StatusCode(500);
            }

        }

        public UserInfoDTO SaveUser(UserInfoDTO user, string coopID)
        {
            try
            {
                if (coopID == null)
                {
                    user.CoopID = "0";
                } else
                {
                    user.CoopID = coopID;
                }
                
                if (user.UserName == null)
                {
                    user.UserName = "NoUserName";
                }
                var output = flexcelsum.SaveUserInfo(user);

                return output;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return null;
                //return StatusCode(500);
            }

        }

        //public UserInfoDTO UpdateUser(UserInfoDTO user)
        //{
        //    try
        //    {
                

        //        if (user.UserName == null)
        //        {
        //            user.UserName = "NoUserName";
        //        }
        //        var output = flexcelsum.UpdateUserInfo(user);

        //        return output;

        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e.InnerException);
        //        return null;
        //        //return StatusCode(500);
        //    }

        //}
    }
}

