using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace AVFCoffeeWebApp.Controllers
{
    public class InputController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}