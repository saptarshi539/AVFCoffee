using CoffeeCore.DTO;
using CoffeeCore.Interfaces;
using Microsoft.AspNetCore.Mvc;
using System;

namespace AVFCoffeeWebApp.APIControllers
{
    [Route("api/[controller]")]
    public class TechnicianHomeAPIController : Controller
    {
        ITechnicianFlexcelSum technianflexcelSum;

        public TechnicianHomeAPIController(ITechnicianFlexcelSum _technianflexcelSum)
        {
            technianflexcelSum = _technianflexcelSum;
        }

        [Route("metrics")]
        [HttpGet]
        [Produces("application/json")]
        public IActionResult GetMetrics()
        {
            try
            {
                TechnicianLoginInfoDTO output = new TechnicianLoginInfoDTO();
                //if (User.Identity.IsAuthenticated)
                //{
                    //var id = User.GetId();
                output = technianflexcelSum.GetUserMetrics();
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

        [Route("savemetrics")]
        [HttpPost]
        [Produces("application/json")]
        public IActionResult SaveMetrics([FromBody]String[] data)
        {
            try
            {
                
               
                //if (User.Identity.IsAuthenticated)
                //{
                //var id = User.GetId();
                technianflexcelSum.saveUserMetrics(data);
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
    }
}
