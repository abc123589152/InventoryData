using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using FactoryData.Controllers;
    
namespace FactoryData.Controllers
{
    
    public class AuthorityController : Controller
    {
        HomeController ho = new HomeController();
        // GET: Authority
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult AuthorityGroup() 
        {
            return View();
        }
        public ActionResult AuthorityDisplay(string SearchAuthority) 
        {
            //List<Authority> list = new List<Authority>();
            //string SelectAuthority = "";
            //if (!string.IsNullOrEmpty(SearchAuthority))
            //{
            //    SelectAuthority += string.Format("Select *from Authority where GroupName ='" + SearchAuthority + "'");
            //}
            //else 
            //{
            //    SelectAuthority += string.Format("Select *from Authority");
            //}
            //DataTable dt = ho.da(SelectAuthority);
            //for (int i = 0; i < dt.Rows.Count; i++) 
            //{
            //    Authority au = new Authority();
            //    au.ID = int.Parse(dt.Rows[i]["ID"].ToString());
            //    au.GroupName = dt.Rows[i]["GroupName"].ToString();
            //    au.PermitList = dt.Rows[i]["PermitList"].ToString();
            //    au.Remarks = dt.Rows[i]["Remarks"].ToString();
            //    list.Add(au);
            //}
            return View();
        }
    }
}