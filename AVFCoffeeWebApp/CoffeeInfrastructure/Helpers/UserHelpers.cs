using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace CoffeeInfrastructure.Helpers
{
    public static class UserHelpers
    {
        public static bool IsUserOwnerOfField(string userID, string fieldID, string connectionStr)
        {
            // query deere orgID and check that the fieldID belongs to the Org
            using (SqlConnection connection = new SqlConnection(connectionStr))
            //Make SQL Connection using AgDBdataReaderConnStr and query database to get emodiisGridded values by pixel, then aggregate
            {
                SqlDataAdapter myCommand = new SqlDataAdapter();

                // Check if user owns the specified field
                string chkOwnFieldQuery = String.Format(@"select COUNT(*) from 
                                                        (
                                                            SELECT [user_id], [deere_org_id], Fields.ID 
                                                            FROM [Farmer_Data].[dbo].[DeereUsers]
                                                            inner join[Farmer_Data].[deere].[Fields] on 
                                                            [Farmer_Data].[dbo].[DeereUsers].deere_org_id =[Farmer_Data].[deere].[Fields].OrganizationID
                                                            where Fields.id = @fieldID  and user_id=@userID
                                                         ) t1");
                connection.Open();
                SqlCommand mycommand = new SqlCommand(chkOwnFieldQuery);
                mycommand.Parameters.Add("@fieldID", SqlDbType.VarChar);
                mycommand.Parameters["@fieldID"].Value = fieldID;
                mycommand.Parameters.Add("@userID", SqlDbType.VarChar);
                mycommand.Parameters["@userID"].Value = userID;
                mycommand.Connection = connection;
                var count = (Int32)mycommand.ExecuteScalar();
                if (count == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }

            }
        }

        public static bool IsUserAuthToAccessOrg(string userID, Int64 organizationID, string connectionStr)
        {
            // query DeereUsers and check that the user have access to OrgID
            using (SqlConnection connection = new SqlConnection(connectionStr))
            //Make SQL Connection
            {
                SqlDataAdapter myCommand = new SqlDataAdapter();

                // Check if user can access the specified orgID
                string chkOwnFieldQuery = String.Format(@"select COUNT(*) from 
                                                        (
                                                            SELECT [deere_org_id]
                                                            FROM [Farmer_Data].[dbo].[DeereUsers]
                                                            WHERE deere_org_id = @orgID AND user_id = @userID
                                                         ) t1");
                connection.Open();
                SqlCommand mycommand = new SqlCommand(chkOwnFieldQuery);
                mycommand.Parameters.Add("@orgID", SqlDbType.VarChar);
                mycommand.Parameters["@orgID"].Value = organizationID;
                mycommand.Parameters.Add("@userID", SqlDbType.VarChar);
                mycommand.Parameters["@userID"].Value = userID;
                mycommand.Connection = connection;
                var count = (Int32)mycommand.ExecuteScalar();
                if (count == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }
    }
}
