using System;
using System.Collections.Generic;

using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using AVFCoffeeWebApp.Models;
using Coffee.APIControllers;
using CoffeeInfrastructure.Helpers;
using Newtonsoft.Json;
using CoffeeCore.DTO;

namespace AVFCoffeeWebApp.Controllers
{
    public class HomeController : Controller
    {
        CellSumController cellSumController;

        public HomeController(CellSumController cellSum)
        {
            cellSumController = cellSum;
        }

        public IActionResult Index()
        {
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

                cellSumController.SaveUser(user);
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
                     return RedirectToAction("", "Results");
                 }
                 else if ( stats == false)
                 {
                     return RedirectToAction("", "Input");
                 }
                 else
                 {
                     return PartialView();
                 }
             }

            else
            {
                return PartialView();
            }
        }
        
     
    }
}
