using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SCM_Win_Service
{
    public class DBUtil
    {
        SqlConnection conn;

        /// <summary>
        /// Initialises the SQL Connection & also opens the Connection
        /// </summary>
        private void init()
        {
            try
            {
                conn = new SqlConnection(ConfigurationManager.ConnectionStrings["VVConnection"].ToString());
                conn.Open();
            }
            catch (Exception ex)
            {
                Logger.LogError(ex, "Exception from connection establisment.");
            }
        }

        public DataSet GetSCMBuyerList()
        {
            DataSet ds = new DataSet();
            try
            {
                this.init();

                SqlCommand cmd = new SqlCommand("[spGetSCMBuyerList]", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = cmd;

                da.Fill(ds);

                conn.Close();
            }
            catch (Exception ex)
            {
                conn.Close();
                Logger.LogError(ex, "Exception from connection establisment.");
            }
            finally
            {
            }

            return ds;
        }


        public DataSet GetSCMBuyerData(string BuyerName)
        {
            DataSet ds = new DataSet();
            try
            {
                this.init();

                SqlCommand cmd = new SqlCommand("[spGetSCMBuyerData]", conn);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add(new SqlParameter("@BuyerName", BuyerName));

                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = cmd;

                da.Fill(ds);

                conn.Close();
            }
            catch (Exception ex)
            {
                conn.Close();
                Logger.LogError(ex, "Exception from connection establisment.");
            }
            finally {
            }

            return ds;
        }
    }
}
