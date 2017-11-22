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
       
        public CellSumController(IFlexcelsum _flexcelsum)
        {
            flexcelsum = _flexcelsum;
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
            bool transition = false, Double workerSalarySoles = 16.16, Double productionQuintales = 14, Double transportCostSoles = 235.22, Double costPriceSolesPerQuintal = 556.51)
        {
            try
            {
                ChartDataDTO sContent = flexcelsum.getOutputFromExcel(earlyHectares,peakHectares,oldHectares,conventional, organic, transition, workerSalarySoles, productionQuintales,
                    transportCostSoles, costPriceSolesPerQuintal);
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
        public IActionResult PostInputs(ChartInputDTO chartInputDTO)
        {
            try
            {
                if (User.Identity.IsAuthenticated)
                {
                    var id = User.GetId(); //"e661c05f-dc88-48c3-8026-3718143c56d8";//
                    flexcelsum.SaveUserInputs(id, chartInputDTO);
                }
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
                if (User.Identity.IsAuthenticated)
                {
                    var id = User.GetId(); //"e661c05f-dc88-48c3-8026-3718143c56d8";//
                    flexcelsum.SaveUserOutputs(id, chartDataDTO);
                }
                //ChartDataDTO sContent = null;
                return Ok();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return StatusCode(500);
            }

        }

        [Route("saveuser")]
        [HttpPost]
        [Produces("application/json")]
        public IActionResult PostUserData([FromBody] UserInfoDTO userInfoDTO)
        {
            try
            {
                if (User.Identity.IsAuthenticated)
                {
                    var id = User.GetId(); //"e661c05f-dc88-48c3-8026-3718143c56d8";//
                    flexcelsum.SaveUserInfo(id, userInfoDTO);
                }
                //ChartDataDTO sContent = null;
                return Ok();
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
        public IActionResult GetInputs(string userid)
        {
            try
            {
                LoginInfoDTO output = new LoginInfoDTO();
                if (User.Identity.IsAuthenticated)
                {
                    var id = User.GetId();
                    //var id = "e661c05f-dc88-48c3-8026-3718143c56d8";
                    output = flexcelsum.GetUserInputs(id);
                }
                //ChartDataDTO sContent = null;
                return Ok(output);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return StatusCode(500);
            }

        }
    }
}

