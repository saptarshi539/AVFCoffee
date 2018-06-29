using CoffeeCore.Interfaces;
using Microsoft.Extensions.Configuration;

namespace CoffeeInfrastructure.FarmerManagement
{
    public class Farmer : IFarmer
    {
        private readonly IConfiguration _iconfiguration;

        public Farmer(IConfiguration configuration)
        {
            _iconfiguration = configuration;
        }

        public bool checkPhoneNumber(string phoneNumber)
        {
            var conn = _iconfiguration.GetSection("ConnectionStrings").GetSection("CoffeeConnStr").Value;
            return false;
        }
    }
}
