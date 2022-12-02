using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using FactoryData.Models;
using System.Data.SqlClient;

namespace FactoryData.Controllers
{
    public class CompanyController : Controller
    {
        SqlConnection sc = new SqlConnection();
        string sql = @"server=192.168.1.36;initial catalog=FactoryData;user id=sa;password=iamthegad123";
        //string sql = @"server=DESKTOP-DGKQ6QE\SQLEXPRESS;initial catalog=FactoryData;user id=sa;password=iamthegad123";

        // GET: Company
        HomeController ho = new HomeController();
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Company(CompanyModel com,string SearchCardNumber)
        {
            string str ;
            if (!string.IsNullOrEmpty(SearchCardNumber))
            {
                str = string.Format("Select *from Company WHERE CardNumber like'%" + SearchCardNumber + "%'");
            }
            else 
            {
                str = string.Format("select *from Company");
            } 
            DataTable dt = ho.da(str);
            return View(dt);
        }
        public ActionResult InsertCompany() 
        {
            return View();
        }
        [HttpPost]
        public ActionResult InsertCompany(CompanyModel com)
        {
           sc.ConnectionString = sql;
            string insertStr = "insert into Company(CompanyName,empolyeesName,CardNumber,AddDatetime)" +
                "VALUES('" + com.CompanyName + "','" + com.empolyeesName + "','" + com.CardNumber + "','" + com.AddDatetime + "')";
            ho.CommandQuerytosql(insertStr);
            return RedirectToAction("Company");
        }
        public ActionResult DeleteCompany(CompanyModel com)
        {
            string Deletestr = string.Format("Delete from Company WHERE ID =" + com.ID.ToString().Replace("'", "''"));
            ho.CommandQuerytosql(Deletestr);
            return RedirectToAction("Company");
        }
        [HttpPost]
        public ActionResult CompanyEdit(CompanyModel edit)
        {
            string SqlEdit = string.Format("UPDATE Company SET CompanyName = '" + edit.CompanyName + "',empolyeesName  ='" + edit.empolyeesName + "',CardNumber ='" + edit.CardNumber+
              "',AddDatetime ='" + edit.AddDatetime + "',EditDatetime = GETDATE()"+" WHERE ID = " + edit.ID);
            ho.CommandQuerytosql(SqlEdit);
            return RedirectToAction("Company");
        }
        public ActionResult CompanyEdit(CompanyEdit comEdit, string id)
        {
            id = comEdit.ID;
            string sqlEditselect = string.Format("Select *from Company WHERE ID = " + comEdit.ID);
            DataTable dt = ho.da(sqlEditselect);
            return View(dt);
        }
    }
}