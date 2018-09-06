using System;
using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;
using Coffee.APIControllers;
using CoffeeInfrastructure.Helpers;
using Newtonsoft.Json;
using CoffeeCore.DTO;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Hosting;

namespace AVFCoffeeWebApp.Controllers
{
    public class HomeController : Controller
    {
        CellSumController cellSumController;
        private readonly IConfiguration _iconfiguration;
        private readonly IHostingEnvironment _env;

        public HomeController(CellSumController cellSum, IConfiguration iconfiguration, IHostingEnvironment env)
        {
            cellSumController = cellSum;
            _iconfiguration = iconfiguration;
            _env = env;
        }

        public IActionResult Index()
        {
            try
            {
                ViewData["apiURL"] = _iconfiguration.GetSection("ProjectVariables").GetSection("apiURL").Value;
                if (User.Identity.IsAuthenticated)
                {
                    var cooperativeID = User.GetCooperativeID();
                    var username = User.GetGivenName();
                    var userID = User.GetId();
                    var language = User.GetSiupinPolicyName();
                    UserInfoDTO user = new UserInfoDTO();
                    user.Language = language;
                    user.UserID = userID;
                    user.UserName = username;

                    cellSumController.SaveUser(user, cooperativeID);
                    //make call to service
                    var inputOutputObject = cellSumController.GetOutputStatus(User.GetId());
                    var outp = inputOutputObject.loginfo["Outputs"];
                    var json = JsonConvert.SerializeObject(outp);
                    Dictionary<String, object> prod = JsonConvert.DeserializeObject<Dictionary<String, object>>(json);
                    var prodOutput = prod["ProducerOutputEnglish"];
                    ProducerOutputEnglishDTO producerEnglish = JsonConvert.DeserializeObject<ProducerOutputEnglishDTO>(prodOutput.ToString());
                    var stats = producerEnglish.status;
                    if (stats == true)
                    {
                        //return RedirectToAction("", "Results");
                        return RedirectToAction("", "TechnicianHome");
                    }
                    else if (stats == false)
                    {
                        return RedirectToAction("", "TechnicianHome");
                    }
                    else
                    {
                        return PartialView();
                    }
                }

                else
                {
                    return PartialView(ViewData["apiURL"]);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return View(e);
            }
        }

    }
}
