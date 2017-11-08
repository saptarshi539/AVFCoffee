using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using CoffeeCore.Interfaces;

namespace Coffee.APIControllers
{
    [Route("api/[controller]/sum")]
    public class CellSumController : Controller
    {
        IFlexcelsum flexcelsum;
       
        public CellSumController(IFlexcelsum _flexcelsum)
        {
            flexcelsum = _flexcelsum;
        }

        [HttpGet]
        public IActionResult CalculateSum()
        {
            try
            {
                
                var sContent = flexcelsum.sumcells();
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

