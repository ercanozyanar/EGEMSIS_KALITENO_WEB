﻿using System;
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
    public class EsnekBaskiController : Controller
    {
        EGEM2021Entities1 db = new EGEM2021Entities1();
        SqlConnection conn1 = new SqlConnection("Data Source=192.168.0.251;Initial Catalog=EGEM2022;Persist Security Info=True;User ID=egem;Password=123456");
        public static string adi;
        public static string fname;
        public static string kod;
        SqlCommand komut = new SqlCommand();


        public ActionResult Baskilist(int sayfa = 1)
        {
            var degerler = db.EGEM_ESNEK_BASKI_KALITE.ToList().ToPagedList(sayfa, 20);
            return View(degerler);
        }
        public ActionResult Baskiduzen(String Search_Data, String Filter_Value, int? Page_No)
        {
            kod = Search_Data;
            if (Search_Data != null)
            {
                Page_No = 1;
            }
            else
            {
                Search_Data = Filter_Value;
            }
            ViewBag.FilterValue = Search_Data;

            var students = from stu in db.EGEM_ESNEK_BASKI_KALITE select stu;
            if (!String.IsNullOrEmpty(Search_Data))
            {
                students = students.Where(stu => stu.ZAMAN.ToString().Contains(Search_Data.ToString()));

                students = students.Where(stu => stu.ZAMAN.ToString().Contains(Search_Data.ToString()));
            }
            students = students.OrderByDescending(stu => stu.ID);
            int Size_Of_Page = 10;
            int No_Of_Page = (Page_No ?? 1);
            return View(students.ToPagedList(No_Of_Page, Size_Of_Page));
        }
        [HttpGet]
        public ActionResult Baskiekle()
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
            ViewBag.KONTROL_ONAY = degerler1;
            
            return View();
        }
        [HttpPost]
        public ActionResult Baskiekle(EGEM_ESNEK_BASKI_KALITE p1)
        {
            List<SelectListItem> degerler1 = (from i in db.EGEM_KALITE_PERSONEL.ToList()
                                              select new SelectListItem
                                              {
                                                  Text = i.ADSOYAD
                                              }).ToList();
            ViewBag.KONTROL_ONAY = degerler1;
            db.EGEM_ESNEK_BASKI_KALITE.Add(p1);
            db.SaveChanges();
            ViewBag.Alert = CommonServices.ShowAlert(Alerts.Success, "Kayıt İşlemi Tamamlanmıştır");
            return View();
        }
        [HttpPost]
        public JsonResult Sipbul(string FISNO)
        {
            EGEM_GRAVUR_KALITE_SIPARIS model = db.EGEM_GRAVUR_KALITE_SIPARIS.FirstOrDefault(f => f.FISNO.Contains(FISNO));
            return Json(model, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult SendEmail(string MSIPARIS_NO, string MBOBINNO,string MHATA,string MIMAJ)
        {
            try
            {
                var senderEmail = new MailAddress("bilgi@egemambalaj.com.tr", "Egemsis");
                var receiverEmail = new MailAddress("ercan.ozyanar@egemambalaj.com.tr");
                var password = "Zug81146";
                var subject = "SIPARIS NO : " + MSIPARIS_NO + " " + " PALET NO :" + MBOBINNO + " " + " Esnek Baski Kalite Hata";
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
                    if (String.IsNullOrEmpty(MIMAJ))
                    {
                        smtp.Send(mess);
                    }
                    else
                    {
                        mess.Attachments.Add(new Attachment(fname));
                        smtp.Send(mess);
                    }
                }
                komut.CommandText = "INSERT INTO EGEM_ESNEK_BASKI_KALITENOKIMAJ (SIPARISNO,BOBINNO,HATA,IMAJ,ADRES) VALUES (@MSIPARIS_NO,@MBOBINNO,@MHATA,@IMAJ,@ADRES)";
                komut.Connection = conn1;
                komut.CommandType = CommandType.Text;
                conn1.Open();
                komut.Parameters.Add("@MSIPARIS_NO", MSIPARIS_NO);
                komut.Parameters.Add("@MBOBINNO", MBOBINNO);
                komut.Parameters.Add("@MHATA", MHATA);
                komut.Parameters.Add("@IMAJ", MIMAJ);
                if (String.IsNullOrEmpty(MIMAJ))
                {
                    komut.Parameters.Add("@ADRES", "");
                }
                else
                {
                    komut.Parameters.Add("@ADRES",adi + "_" + MIMAJ);
                }
                komut.ExecuteReader();
                conn1.Close();
                return View();
            }
            catch (Exception)
           {
               return View();
           }
        }
        public ActionResult Sil(int id)
        {
            var baski = db.EGEM_ESNEK_BASKI_KALITE.Find(id);
            db.EGEM_ESNEK_BASKI_KALITE.Remove(baski);
            db.SaveChanges();
            return RedirectToAction("Baskiduzen");
        }
        [HttpPost]
        public JsonResult DosyaYukle()
        {
            if (Request.Files.Count > 0)
            {
                try
                {
                    HttpFileCollectionBase files = Request.Files;
                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];
                        //string fname;
                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            fname = testfiles[testfiles.Length - 1];
                        }
                        else
                        {
                            string dosyaadi = DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();
                            string fileName = System.IO.Path.GetFileName(file.FileName);
                            fname = Path.Combine(Server.MapPath("~/Content/NOK_IMAJ/ESNEKBASKI/"), dosyaadi +"_"+ fileName);
                            file.SaveAs(fname);
                            adi = dosyaadi;
                        }
                      }
                    return Json("Dosya Yükleme Başarılı");
                }
                catch (Exception ex)
                {
                    return Json("Hata Oluştu:  " + ex.Message);
                }
            }
            else
            {
                return Json("Dosya seçilmedi");
            }
        }

        public ActionResult BaskiNOK(String Search_Data, String Filter_Value, int? Page_No)
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

            var students = from stu in db.EGEM_ESNEK_BASKI_KALITENOKIMAJ select stu;
            if (!String.IsNullOrEmpty(Search_Data))
            {
                students = students.Where(stu => stu.SIPARISNO.ToString().Contains(Search_Data.ToString()));

                students = students.Where(stu => stu.SIPARISNO.ToString().Contains(Search_Data.ToString()));
            }
            students = students.OrderByDescending(stu => stu.ID);
            int Size_Of_Page = 10;
            int No_Of_Page = (Page_No ?? 1);
            return View(students.ToPagedList(No_Of_Page, Size_Of_Page));
        }
        [HttpPost]
        public ActionResult Export(string Search_Data, string Filter_Value)
        {
             EGEM2021Entities1 entities = new EGEM2021Entities1();
             DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[34] 
            {
                new DataColumn("URUN_KODU"),
                new DataColumn("URUN_ADI"),
                new DataColumn("SIPARIS_NO"),
                new DataColumn("ZAMAN"),
                new DataColumn("BOBIN_NO"),
                new DataColumn("HAMMADDE_KONTROL"),
                new DataColumn("BOBIN_ENI"),
                new DataColumn("FOTOSEL_ARALIGI"),
                new DataColumn("NET_BASKI_ENI"),
                new DataColumn("BASKI_YONU"),
                new DataColumn("FILM_TIPI"),
                new DataColumn("BARKOD_OKUNUR"),
                new DataColumn("ACILIS_YONU"),
                new DataColumn("ARKA_KAPANMA_SEKLI"),
                new DataColumn("ICERIK_KONTROL"),
                new DataColumn("RENK_UYGUN"),
                new DataColumn("RENK_UYGUN_GOZ"),
                new DataColumn("BASKI_KAYIK"),
                new DataColumn("GORSEL_KONTROL"),
                new DataColumn("ADEZYON"),
                new DataColumn("KURUMA"),
                new DataColumn("TONLAMA"),
                new DataColumn("SAKAL"),
                new DataColumn("CIZGI"),
                new DataColumn("ISIL_YAPISMA"),
                new DataColumn("MUREKKEP_MIKTAR"),
                new DataColumn("BIRIM_AGIRLIK"),
                new DataColumn("COF"),
                new DataColumn("GC_TEST"),
                new DataColumn("BOBIN_ICBAYRAK_SAYI"),
                new DataColumn("MAKINE"),
                new DataColumn("KONTROL_ONAY"),
                new DataColumn("SILINDIR_KOD_KONTROL"),
                new DataColumn("GLOSS")
            });
            var customers = from stu in db.EGEM_ESNEK_BASKI_KALITE select stu;
            customers = customers.Where(stu => stu.ZAMAN.ToString().Contains(kod.ToString()));
            foreach (var customer in customers.ToList())
            {
                dt.Rows.Add(
                    customer.URUN_KODU, 
                    customer.URUN_ADI, 
                    customer.SIPARIS_NO, 
                    customer.ZAMAN, 
                    customer.BOBIN_NO,
                    customer.HAMMADDE_KONTROL,
                    customer.BOBIN_ENI,
                    customer.FOTOSEL_ARALIGI,
                    customer.NET_BASKI_ENI,
                    customer.BASKI_YONU,
                    customer.FILM_TIPI,
                    customer.BARKOD_OKUNUR,
                    customer.ACILIS_YONU,
                    customer.ARKA_KAPANMA_SEKLI,
                    customer.ICERIK_KONTROL,
                    customer.RENK_UYGUN,
                    customer.RENK_UYGUN_GOZ,
                    customer.BASKI_KAYIK,
                    customer.GORSEL_KONTROL,
                    customer.ADEZYON,
                    customer.KURUMA,
                    customer.TONLAMA,
                    customer.SAKAL,
                    customer.CIZGI,
                    customer.ISIL_YAPISMA,
                    customer.MUREKKEP_MIKTAR,
                    customer.BIRIM_AGIRLIK,
                    customer.COF,
                    customer.GC_TEST,
                    customer.BOBIN_ICBAYRAK_SAYI,
                    customer.MAKINE,
                    customer.KONTROL_ONAY,
                    customer.SILINDIR_KOD_KONTROL,
                    customer.GLOSS
                    );
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
