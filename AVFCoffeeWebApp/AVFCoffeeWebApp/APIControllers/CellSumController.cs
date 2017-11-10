using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;

using CoffeeCore.Interfaces;
using CoffeeInfrastructure.Flexcel;

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

        [Route("sum/{cellId:long}")]
        [HttpGet]
        public IActionResult CalculateSum(long cellId)
        {
            try
            {
                var l = cellId;
                String sContent = flexcelsum.sumcells();
                return Ok(sContent);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return StatusCode(500);
            }

        }

        [Route("calculate/earlyHectares/{earlyHectares}/peakHectares/{peakHectares}/oldHectares/{oldHectares}/conventional/{conventional}/organic/{organic}/transition/{transition}/workerSalarySoles" +
            "/{workerSalarySoles}/productionQuintales/{productionQuintales}/transportCostSoles/{transportCostSoles}/costPriceSolesPerQuintal/{costPriceSolesPerQuintal}")]
        [HttpGet]
        [Produces("application/json")]
        public IActionResult GetChartValues(Double earlyHectares = 0.2, Double peakHectares = 0.31, Double oldHectares = 0.37, bool conventional = true, bool organic = true, 
            bool transition = true, Double workerSalarySoles = 225, Double productionQuintales = 23, Double transportCostSoles = 235, Double costPriceSolesPerQuintal = 256)
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
    }
}

