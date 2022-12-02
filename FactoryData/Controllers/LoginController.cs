using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using FactoryData.Models;
using System.Globalization;
using System.Data.SqlClient;
using System.Security.Cryptography;
using System.Text;
namespace FactoryData.Controllers
{
    public class LoginController : Controller
    {
        public ActionResult Login() 
        {
            return View(); ;
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Logout() {
            FormsAuthentication.SignOut();
            Session.Abandon();
            // clear authentication cookie
            //HttpCookie cookie1 = new HttpCookie(FormsAuthentication.FormsCookieName, "");
            //cookie1.Expires = DateTime.Now.AddYears(-1);
            //Response.Cookies.Add(cookie1);           
            //FormsAuthentication.RedirectToLoginPage(); 
            return RedirectToAction("Index", "Home", null);
        }
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Login(Login login)
        {
            FactoryDataEntities2 db= new FactoryDataEntities2();// 登入的密碼（以 SHA1 加密）   

            //這一條是去資料庫抓取輸入的帳號密碼的方法請自行實做
            //login.UserPassWord = FormsAuthentication.HashPasswordForStoringInConfigFile(login.UserPassWord, "SHA1");
            //SHA1CryptoServiceProvider sha1 = new SHA1CryptoServiceProvider();
            
            //byte[] password_bytes = Encoding.ASCII.GetBytes(login.UserPassWord);
            //byte[] encrypt_bytes = sha1.ComputeHash(password_bytes);
            //string loginpass = Convert.ToString(encrypt_bytes);            
            var r = db.Account.Where(x => x.UserName == login.UserName && x.UserPassWord == login.UserPassWord).FirstOrDefault();
            
                if (r == null)
                {
                    TempData["Error"] = "您輸入的帳號不存在或者密碼錯誤!";
                    return View();
                }
                else if (r != null)
                {
                    Session["UserID"] = r.ID.ToString();
                    Session["Remarks"] = r.Remarks.ToString();
                    Session["UserName"] = r.UserName.ToString();
                    Session["Permit"] = r.Permit.ToString();
                    
                }
           
            // 登入時清空所有 Session 資料
            //Session.RemoveAll();
            FormsAuthenticationTicket ticket = new FormsAuthenticationTicket(
              r.ID,
              r.UserName,//你想要存放在 User.Identy.Name 的值，通常是使用者帳號
              DateTime.Now,
              DateTime.Now.AddMinutes(30),
              false,//將管理者登入的 Cookie 設定成 Session Cookie
              r.Remarks.ToString(),//userdata看你想存放甚麼資訊
              FormsAuthentication.FormsCookiePath);
            string encTicket = FormsAuthentication.Encrypt(ticket);
            Response.Cookies.Add(new HttpCookie(FormsAuthentication.FormsCookieName,encTicket));
            return RedirectToAction("Index", "Home");
        }
        // GET: Account

        //public ActionResult Login(Login model)
        //{
        //FactoryDataEntities db = new FactoryDataEntities();
        //if (!ModelState.IsValid)
        //{
        //    return View(model);
        //}

        //var user = db.Account.Where(x => x.UserName == model.UserName && x.UserPassWord == model.UserPassWord).FirstOrDefault();

        //if (user == null)
        //{
        //    ModelState.AddModelError("", "無效的帳號或密碼。");
        //    return View();
        //}

        //var ticket = new FormsAuthenticationTicket(
        //            version: 1,
        //            name: user.ID.ToString(), //可以放使用者Id
        //            issueDate: DateTime.UtcNow,//現在UTC時間
        //            expiration: DateTime.UtcNow.AddMinutes(30),//Cookie有效時間=現在時間往後+30分鐘
        //            isPersistent: true,// 是否要記住我 true or false
        //            userData, //可以放使用者角色名稱
        //            cookiePath: FormsAuthentication.FormsCookiePath);

        //var encryptedTicket = FormsAuthentication.Encrypt(ticket); //把驗證的表單加密
        //var cookie = new HttpCookie(FormsAuthentication.FormsCookieName, encryptedTicket);
        //Response.Cookies.Add(cookie);
        //    return RedirectToAction("Index", "Home");
        //}
    }
}