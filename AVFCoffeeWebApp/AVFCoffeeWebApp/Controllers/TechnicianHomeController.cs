using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

namespace AVFCoffeeWebApp.Controllers
{
    public class TechnicianHomeController : Controller
    {
        private readonly IConfiguration _iconfiguration;
        private readonly IHostingEnvironment _env;

        public TechnicianHomeController(IConfiguration iconfiguration, IHostingEnvironment env)
        {
            _iconfiguration = iconfiguration;
            _env = env;
        }
        [Authorize]
        public IActionResult Index()
        {
            ViewData["apiURL"] = _iconfiguration.GetSection("ProjectVariables").GetSection("apiURL").Value;
            return View();
        }
    }
}
