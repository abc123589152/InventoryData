using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
namespace FactoryData.Models
{
    public class ReturnPermit
    {
        public string returnDoorper(string pertext,string sql) 
        {         
            string sqltest = "Select *from Permition where GroupName = '" + pertext+"'";
            SqlDataAdapter sd = new SqlDataAdapter(sqltest,sql);
            DataSet ds = new DataSet();
            sd.Fill(ds);
            DataTable data = ds.Tables[0];
            string ReturnTrueOrFalse = data.Rows[0]["DOOR"].ToString();
            return ReturnTrueOrFalse;
        }
        public string returnPermitDevice(string pertext, string sql)
        {
            string perselect = "Select *from Permition where GroupName = '" + pertext + "'";
            SqlDataAdapter sd = new SqlDataAdapter(perselect, sql);
            DataSet ds = new DataSet();
            sd.Fill(ds);
            DataTable PerDt = ds.Tables[0];
            string dtText = PerDt.Rows[0]["Device"].ToString();
            return dtText;
        }
        public string returnPermitFab(string pertext,string sql)
        {
            string perselect = "Select *from Permition where GroupName = '" + pertext + "'";
            SqlDataAdapter sd = new SqlDataAdapter(perselect, sql);
            DataSet ds = new DataSet();
            sd.Fill(ds);
            DataTable PerDt = ds.Tables[0];
            string dtText = PerDt.Rows[0]["Fab"].ToString();
            return dtText;
        }
        public string returnPermitCCTV(string pertext,string sql)
        {
            string perselect = "Select *from Permition where GroupName = '" + pertext + "'";
            SqlDataAdapter sd = new SqlDataAdapter(perselect, sql);
            DataSet ds = new DataSet();
            sd.Fill(ds);
            DataTable PerDt = ds.Tables[0];
            string dtText = PerDt.Rows[0]["CCTV"].ToString();
            return dtText;
        }
        public string returnPermitReport(string pertext, string sql)
        {
            string perselect = "Select *from Permition where GroupName = '" + pertext + "'";
            SqlDataAdapter sd = new SqlDataAdapter(perselect, sql);
            DataSet ds = new DataSet();
            sd.Fill(ds);
            DataTable PerDt = ds.Tables[0];
            string dtText = PerDt.Rows[0]["Report"].ToString();
            return dtText;
        }
        public string returnPermitOrnerReport(string pertext, string sql)
        {
            string perselect = "Select *from Permition where GroupName = '" + pertext + "'";
            SqlDataAdapter sd = new SqlDataAdapter(perselect, sql);
            DataSet ds = new DataSet();
            sd.Fill(ds);
            DataTable PerDt = ds.Tables[0];
            string dtText = PerDt.Rows[0]["OrnerReport"].ToString();
            return dtText;
        }
        public string returnPermitAccount(string pertext, string sql)
        {
            string perselect = "Select *from Permition where GroupName = '" + pertext + "'";
            SqlDataAdapter sd = new SqlDataAdapter(perselect, sql);
            DataSet ds = new DataSet();
            sd.Fill(ds);
            DataTable PerDt = ds.Tables[0];
            string dtText = PerDt.Rows[0]["Account"].ToString();
            return dtText;
        }
        public string returnPermitPermit(string pertext, string sql)
        {
            string perselect = "Select *from Permition where GroupName = '" + pertext + "'";
            SqlDataAdapter sd = new SqlDataAdapter(perselect, sql);
            DataSet ds = new DataSet();
            sd.Fill(ds);
            DataTable PerDt = ds.Tables[0];
            string dtText = PerDt.Rows[0]["Permit"].ToString();
            return dtText;
        }
        public string returnSMSCCTV(string pertext, string sql)
        {
            string perselect = "Select *from Permition where GroupName = '" + pertext + "'";
            SqlDataAdapter sd = new SqlDataAdapter(perselect, sql);
            DataSet ds = new DataSet();
            sd.Fill(ds);
            DataTable PerDt = ds.Tables[0];
            string dtText = PerDt.Rows[0]["SMSCCTV"].ToString();
            return dtText;
        }
        public string returnSMSReader(string pertext, string sql)
        {
            string perselect = "Select *from Permition where GroupName = '" + pertext + "'";
            SqlDataAdapter sd = new SqlDataAdapter(perselect, sql);
            DataSet ds = new DataSet();
            sd.Fill(ds);
            DataTable PerDt = ds.Tables[0];
            string dtText = PerDt.Rows[0]["SMSReader"].ToString();
            return dtText;
        }       
        public string returnSMSAlarmSystem(string pertext, string sql)
        {
            string perselect = "Select *from Permition where GroupName = '" + pertext + "'";
            SqlDataAdapter sd = new SqlDataAdapter(perselect, sql);
            DataSet ds = new DataSet();
            sd.Fill(ds);
            DataTable PerDt = ds.Tables[0];
            string dtText = PerDt.Rows[0]["SMSAlarmSystem"].ToString();
            return dtText;
        }
        public string returnSMSBarriergate(string pertext, string sql)
        {
            string perselect = "Select *from Permition where GroupName = '" + pertext + "'";
            SqlDataAdapter sd = new SqlDataAdapter(perselect, sql);
            DataSet ds = new DataSet();
            sd.Fill(ds);
            DataTable PerDt = ds.Tables[0];
            string dtText = PerDt.Rows[0]["SMSBarriergate"].ToString();
            return dtText;
        }
        public string returnSMSGate(string pertext, string sql)
        {
            string perselect = "Select *from Permition where GroupName = '" + pertext + "'";
            SqlDataAdapter sd = new SqlDataAdapter(perselect, sql);
            DataSet ds = new DataSet();
            sd.Fill(ds);
            DataTable PerDt = ds.Tables[0];
            string dtText = PerDt.Rows[0]["SMSGate"].ToString();
            return dtText;
        }
        public string returnSMSMakeCard(string pertext, string sql)
        {
            string perselect = "Select *from Permition where GroupName = '" + pertext + "'";
            SqlDataAdapter sd = new SqlDataAdapter(perselect, sql);
            DataSet ds = new DataSet();
            sd.Fill(ds);
            DataTable PerDt = ds.Tables[0];
            string dtText = PerDt.Rows[0]["SMSMakeCard"].ToString();
            return dtText;
        }
    }
}