using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

namespace AVFCoffeeWebApp.Controllers
{
    public class AdvancedInputsController : Controller
    {
        private readonly IConfiguration _iconfiguration;
        private readonly IHostingEnvironment _env;

        public AdvancedInputsController(IConfiguration iconfiguration, IHostingEnvironment env)
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
