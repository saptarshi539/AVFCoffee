using Domain.Flexcel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Threading.Tasks;

namespace CofeeWebAPI.Controllers
{
    [Route("api/[controller]")]
    public class CellSumController :Controller
    {

        private readonly IConfiguration _iconfiguration;
        private readonly IHostingEnvironment _env;
        internal string _rmaRatesConn;
        public CellSumController(IConfiguration iconfiguration, IHostingEnvironment env)
        {
            _iconfiguration = iconfiguration;
            _env = env;
            _rmaRatesConn = _iconfiguration.GetSection("ConnectionStrings").GetSection("RMAdataReaderConnStr").Value;
        }

        [Route("sum")]
        [HttpPost]
        public async Task<IActionResult> CalculateSum()
        {
            try
            {
                Flexcelsum fsum = new Flexcelsum();
                var sContent = fsum.sumcells();
                return Ok(sContent);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.InnerException);
                return StatusCode(500);
            }

        }
    }
}

