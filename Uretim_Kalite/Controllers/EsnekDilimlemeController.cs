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

namespace Uretim_Kalite.Controllers
{
    public class EsnekDilimlemeController : Controller
    {
        EGEM2021Entities1 db = new EGEM2021Entities1();
        SqlConnection conn1 = new SqlConnection("Data Source=192.168.0.251;Initial Catalog=EGEM2021;Persist Security Info=True;User ID=egem;Password=123456");

        public ActionResult Dilimlemelist(int sayfa = 1)
        {
            var degerler = db.EGEM_ESNEK_DILIMLEME_KALITE.ToList().ToPagedList(sayfa, 20);
            return View(degerler);
        }

        public ActionResult Dilimlemeduzen(string Search_Data, string Filter_Value, int? Page_No)
        {
            if (Search_Data != null)
            {
                Page_No = 1;
            }
            else
            {
                Search_Data = Filter_Value;
            }

            ViewBag.FilterValue = Search_Data;

            var students = from stu in db.EGEM_ESNEK_DILIMLEME_KALITE select stu;

            if (!String.IsNullOrEmpty(Search_Data))
            {
                students = students.Where(stu => stu.SIPARIS_NO.ToUpper().Contains(Search_Data.ToUpper()));
            }

            students = students.OrderByDescending(stu => stu.ID);




            int Size_Of_Page = 10;
            int No_Of_Page = (Page_No ?? 1);
            return View(students.ToPagedList(No_Of_Page, Size_Of_Page));
        }

        [HttpGet]
        public ActionResult Dilimlemeekle()
        {
            List<SelectListItem> degerler = (from i in db.EGEM_GRAVUR_KALITE_SIPARIS.ToList()
                                             select new SelectListItem
                                             {
                                                 Text = i.FISNO
                                             }).ToList();
            ViewBag.dgr = degerler;
            List<SelectListItem> degerler1 = (from i in db.EGEM_KALITE_PERSONEL.ToList()
                                              select new SelectListItem
                                              {
                                                  Text = i.ADSOYAD
                                              }).ToList();
            ViewBag.KALITE_ONAY = degerler1;
            return View();
        }
        [HttpPost]
        public ActionResult Dilimlemeekle(EGEM_ESNEK_DILIMLEME_KALITE p1)
        {
            List<SelectListItem> degerler1 = (from i in db.EGEM_KALITE_PERSONEL.ToList()
                                              select new SelectListItem
                                              {
                                                  Text = i.ADSOYAD
                                              }).ToList();
            ViewBag.KALITE_ONAY = degerler1;

            db.EGEM_ESNEK_DILIMLEME_KALITE.Add(p1);
            db.SaveChanges();
            return View();
        }
        [HttpPost]
        public JsonResult Sipbul(string FISNO)
        {
            EGEM_GRAVUR_KALITE_SIPARIS model = db.EGEM_GRAVUR_KALITE_SIPARIS.FirstOrDefault(f => f.FISNO.Contains(FISNO));
            return Json(model, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult SendEmail(string MSIPARIS_NO, string MBOBINNO, string MHATA)
        {
            try
            {
                var senderEmail = new MailAddress("bilgi@egemambalaj.com.tr", "Egemsis");
                var receiverEmail = new MailAddress("ercan.ozyanar@egemambalaj.com.tr,egemsis.kalite@egemambalaj.com.tr", "Receiver");
                var password = "Zug81146";
                var subject = "SIPARIS NO : " + MSIPARIS_NO + " " + " PALET NO :" + MBOBINNO + " " + " Esnek Dilimleme Kalite Hata";
                var body = "SIPARIS NO : " + MSIPARIS_NO + " " + " PALET NO :" + MBOBINNO + " " + " TARIH :" + DateTime.Now + " HATA BILDIRIMI :" + MHATA;
                var smtp = new SmtpClient
                {
                    Host = "smtp.office365.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(senderEmail.Address, password)
                };
                using (var mess = new MailMessage(senderEmail, receiverEmail)
                {
                    Subject = subject,
                    Body = body
                })
                {
                    smtp.Send(mess);
                }


                return View();

            }
            catch (Exception)
            {

                return View();
            }
                               }
        public ActionResult Sil(int id)
        {
            var dilimleme = db.EGEM_ESNEK_DILIMLEME_KALITE.Find(id);
            db.EGEM_ESNEK_DILIMLEME_KALITE.Remove(dilimleme);
            db.SaveChanges();
            return RedirectToAction("Dilimlemeduzen");

        }
        [HttpPost]
        public FileResult Export()
        {
            EGEM2021Entities1 entities = new EGEM2021Entities1();
            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[5] { new DataColumn("ZAMAN"),
                                            new DataColumn("SIPARIS_NO"),
                                            new DataColumn("URUN_KODU"),
                                            new DataColumn("URUN_ADI"),
                                            new DataColumn("TONLAMA")});

            var customers = from customer in entities.EGEM_ESNEK_DILIMLEME_KALITE.Take(10)
                            select customer;

            foreach (var customer in customers)
            {
                dt.Rows.Add(customer.ZAMAN, customer.SIPARIS_NO, customer.URUN_KODU, customer.URUN_ADI, customer.TONLAMA);
            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Grid.xlsx");
                }
            }
        }
    }
}