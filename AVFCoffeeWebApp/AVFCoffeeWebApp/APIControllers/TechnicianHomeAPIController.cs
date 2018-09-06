using CoffeeCore.DTO;
using CoffeeCore.Interfaces;
using Microsoft.AspNetCore.Mvc;
using CoffeeInfrastructure.Helpers;
using System;
using System.Collections.Generic;

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

        [Route("getinputs")]
        [HttpGet]
        [Produces("application/json")]
        public IActionResult GetInputs(string language)
        {
            try
            {

                var inputs = technianflexcelSum.getInputs(language);
                
                return Ok(inputs);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return StatusCode(500);
            }

        }

        [Route("getinputvalues")]
        [HttpGet]
        [Produces("application/json")]
        public IActionResult GetInputValues()
        {
            try
            {

                var inputs = technianflexcelSum.getInputValues();

                return Ok(inputs);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return StatusCode(500);
            }

        }

        [Route("saveinputvalues")]
        [HttpPost]
        [Produces("application/json")]
        public IActionResult SaveInputValues([FromBody]ChartInputAdvancedDTO advancedInputsValues)
        {
            try
            {

                technianflexcelSum.saveUserAdvancedInputs(advancedInputsValues);

                return Ok();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return StatusCode(500);
            }

        }

        [Route("GetAnalysis")]
        [HttpGet]
        [Produces("application/json")]
        public IActionResult GetAnalyses()
        {
            try
            {
                var analyses = new Dictionary<string, List<AnalysisDTO>>();
                //get userid
                if (User.Identity.IsAuthenticated)
                {
                    var userid = User.GetId();
                    var analysesList = technianflexcelSum.GetAnalysis(userid);
                    analyses.Add("analyses", analysesList);
                }
                return Ok(analyses);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return StatusCode(500);
            }

        }

        [Route("GetFarms")]
        [HttpGet]
        [Produces("application/json")]
        public IActionResult GetFarms()
        {
            try
            {
                var farms = new List<FarmInfoDTO>();
                //get coopid
                if (User.Identity.IsAuthenticated)
                {
                    var coopid = User.GetCooperativeID();
                    farms = technianflexcelSum.GetFarms(coopid);
                }
                return Ok(farms);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return StatusCode(500);
            }

        }
    }
}
