using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using AVFCoffeeWebApp.Models;
using Coffee.APIControllers;
using CoffeeInfrastructure.Helpers;

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
                //make call to service
                var inputOutputObject = cellSumController.GetOutputStatus(User.GetId());

                return RedirectToAction("", "Input");
            }
            else
            {
                return PartialView();
            }
        }
        
     
    }
}
