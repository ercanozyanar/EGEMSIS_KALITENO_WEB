using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Uretim_Kalite.Models.Entity;
using System.Data.SqlClient;
using System.Web.Security;
using Uretim_Kalite.Enums;
using Uretim_Kalite.Services;

namespace Uretim_Kalite.Controllers
{
    [AllowAnonymous]
    public class LoginController : Controller
    {
        EGEM2021Entities1 db = new EGEM2021Entities1();
        SqlConnection conn1 = new SqlConnection("Data Source=192.168.0.251;Initial Catalog=EGEM2021;Persist Security Info=True;User ID=egem;Password=123456");

        [Authorize]
        public ActionResult Login()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Login(EGEM_KALITE_USER user)


        {
            var userIndb = db.EGEM_KALITE_USER.FirstOrDefault(x => x.AD_SOYAD == user.AD_SOYAD && x.SIFRE == user.SIFRE);
            DateTime tarih = DateTime.Today;
            DateTime ltarih = Convert.ToDateTime("31.12.2022");
            if (tarih < ltarih)
            {
                if (userIndb != null)
                {
                    ViewBag.Alert = CommonServices.ShowAlert(Alerts.Success, "Employee added");
                    FormsAuthentication.SetAuthCookie(user.AD_SOYAD, false);
                    return RedirectToAction("Baskiekle", "EsnekBaski");
                }
                else
                {
                    ViewBag.Alert = CommonServices.ShowAlert(Alerts.Danger, "Kullanıcı Adı veya Şifre Hatalı...");
                    return View();
                }
            }
            else
            {
                ViewBag.Alert = CommonServices.ShowAlert(Alerts.Danger, "Kullanıcı Adı veya Şifre Hatalı...");
                return View();
            }
        }
    }
}