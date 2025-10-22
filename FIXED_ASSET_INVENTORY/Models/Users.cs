using Microsoft.Data.SqlClient;

namespace FIXED_ASSET_INVENTORY.Models
{
    public class Users
    { 
        public static string userExists(string username, string _connStr)
        {
            SqlConnection con = new SqlConnection(_connStr);
            string result = "You don't have access to this App.";
            try
            {
                con.Open();
                SqlCommand cmd = new SqlCommand("SELECT * FROM [BASE].[dbo].[PD_ACCESS] WHERE ID_USER =@user AND PROGRAM = 'FIXED ASSET INVENTORY'", con);
                cmd.Parameters.AddWithValue("@user", username);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result= "OK";
                }
            }
            catch (Exception ex)
            {
                result = "Could not verify";
            }
            return result;
        }
    }
}
