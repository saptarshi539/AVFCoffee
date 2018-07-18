using CoffeeCore.Interfaces;
using Microsoft.Extensions.Configuration;
using System;
using System.Data.SqlClient;

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
            using (SqlConnection con = new SqlConnection(conn))
            {
                con.Open();
                SqlCommand comm = new SqlCommand("Select * from [AVFCoffee].[dbo].[SmallHolder] where PhoneNumber = @phone", con);
                comm.Parameters.AddWithValue("@phone", phoneNumber);
                // int result = command.ExecuteNonQuery();
                using (SqlDataReader reader = comm.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return true;
                    }
                }
            }
            return false;
        }
    }
}
