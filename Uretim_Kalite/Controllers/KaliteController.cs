using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Uretim_Kalite.Models.Entity;
using PagedList.Mvc;
using PagedList;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Net;
using System.Data;
using ClosedXML.Excel;
using System.IO;
using System.Web.Security;
using Uretim_Kalite.Enums;
using Uretim_Kalite.Services;

namespace Uretim_Kalite.Controllers
{
    public class KaliteController : Controller
    {
        EGEM2021Entities1 db = new EGEM2021Entities1();
        SqlConnection conn1 = new SqlConnection("Data Source=192.168.0.251;Initial Catalog=EGEM2021;Persist Security Info=True;User ID=egem;Password=123456");


        public ActionResult Kalitelist(int sayfa = 1)
        {
            var degerler = db.EGEM_KALITE_PERSONEL.ToList().ToPagedList(sayfa, 20);
            return View(degerler);
        }

        [HttpGet]
        public ActionResult Kaliteekle()
        {
           return View();
        }
        [HttpPost]
        public ActionResult Kaliteekle(EGEM_KALITE_PERSONEL p3)
        {
            db.EGEM_KALITE_PERSONEL.Add(p3);
            db.SaveChanges();
            return View();
        }
        public ActionResult Sil(int id)
        {
            var baski = db.EGEM_KALITE_PERSONEL.Find(id);
            db.EGEM_KALITE_PERSONEL.Remove(baski);
            db.SaveChanges();
            return RedirectToAction("Kalitelist");

        }
    }
}