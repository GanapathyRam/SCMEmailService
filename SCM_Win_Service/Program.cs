using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SCM_Win_Service
{
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                DataSet ds = new DataSet();
                string buyerName, email = string.Empty;

                ds = GetBuyerList();

                for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                {
                    DataSet dsForBuyerData = new DataSet();

                    buyerName = ds.Tables[0].Rows[i]["BuyerName"].ToString();
                    email = ds.Tables[0].Rows[i]["Email"].ToString();

                    if (!string.IsNullOrEmpty(email))
                    {
                        dsForBuyerData = GetBuyerData(buyerName);

                        if (dsForBuyerData.Tables[0].Rows.Count > 0)
                        {
                           Email.ExportDataSet(dsForBuyerData, email, buyerName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError(ex, "Exception from fetching data from database");
            }
        }

        private static DataSet GetBuyerList()
        {
            DataSet ds = new DataSet();
            try
            {
                DBUtil _dbObj = new DBUtil();

                ds = _dbObj.GetSCMBuyerList();
            }
            catch (Exception ex)
            {
                Logger.LogError(ex, "Exception from GetBuyerList()");
            }

            return ds;
        }

        private static DataSet GetBuyerData(string Name)
        {
            DataSet ds = new DataSet();
            try
            {
                DBUtil _dbObj = new DBUtil();

                ds = _dbObj.GetSCMBuyerData(Name);
            }
            catch (Exception ex)
            {
                Logger.LogError(ex, "Exception from GetBuyerData()");
            }

            return ds;
        }
    }
}
