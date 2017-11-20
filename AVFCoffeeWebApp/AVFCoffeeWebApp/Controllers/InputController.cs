using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Hosting;

namespace AVFCoffeeWebApp.Controllers
{
    public class InputController : Controller
    {
        private readonly IConfiguration _iconfiguration;
        private readonly IHostingEnvironment _env;

        public InputController(IConfiguration iconfiguration, IHostingEnvironment env)
        {
            _iconfiguration = iconfiguration;
            _env = env;
        }

        public IActionResult Index()
        {
            ViewData["apiURL"] = _iconfiguration.GetSection("ProjectVariables").GetSection("apiURL").Value;
            return View();
        }
    }
}