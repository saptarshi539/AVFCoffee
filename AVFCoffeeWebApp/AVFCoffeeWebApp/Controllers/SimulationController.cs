using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Authorization;

// For more information on enabling MVC for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace AVFCoffeeWebApp.Controllers
{
    public class SimulationController : Controller
    {
        private readonly IConfiguration _iconfiguration;
        private readonly IHostingEnvironment _env;

        public SimulationController(IConfiguration iconfiguration, IHostingEnvironment env)
        {
            _iconfiguration = iconfiguration;
            _env = env;
        }

        // GET: /<controller>/
        [Authorize]
        public IActionResult Index()
        {
            ViewData["apiURL"] = _iconfiguration.GetSection("ProjectVariables").GetSection("apiURL").Value;
            return View();
        }
    }
}
