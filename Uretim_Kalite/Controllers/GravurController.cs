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
    public class GravurController : Controller
    {
        EGEM2021Entities1 db = new EGEM2021Entities1();
        SqlConnection conn1 = new SqlConnection("Data Source=192.168.0.251;Initial Catalog=EGEM2022;Persist Security Info=True;User ID=egem;Password=123456");
        public static string adi;
        public static string fname;
        SqlCommand komut = new SqlCommand();


        public ActionResult Gravurlist(int sayfa = 1)
        {
            var degerler = db.EGEM_GRAVUR_KALITE.ToList().ToPagedList(sayfa, 20);
            return View(degerler);
        }
        public ActionResult Gravurduzen(string Search_Data, string Filter_Value, int? Page_No)
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

            var students = from stu in db.EGEM_GRAVUR_KALITE select stu;

            if (!String.IsNullOrEmpty(Search_Data))
            {
                students = students.Where(stu => stu.sipno.ToUpper().Contains(Search_Data.ToUpper()));
            }

            students = students.OrderByDescending(stu => stu.ID);

            int Size_Of_Page = 10;
            int No_Of_Page = (Page_No ?? 1);
            return View(students.ToPagedList(No_Of_Page, Size_Of_Page));
        }
        [HttpGet]
        public ActionResult Gravurekle()
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
            ViewBag.kontrol_onay = degerler1;
            return View();
        }
        [HttpPost]
        public ActionResult Gravurekle(EGEM_GRAVUR_KALITE p1)
        {
            List<SelectListItem> degerler1 = (from i in db.EGEM_KALITE_PERSONEL.ToList()
                                              select new SelectListItem
                                              {
                                                  Text = i.ADSOYAD
                                              }).ToList();
            ViewBag.kontrol_onay = degerler1;

            db.EGEM_GRAVUR_KALITE.Add(p1);
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
        public ActionResult SendEmail(string MSIPARIS_NO, string MBOBINNO, string MHATA, string MIMAJ)
        {
            try
            {
                var senderEmail = new MailAddress("bilgi@egemambalaj.com.tr", "Egemsis");
                var receiverEmail = new MailAddress("ercan.ozyanar@egemambalaj.com.tr");
                var password = "Zug81146";
                var subject = "SIPARIS NO : " + MSIPARIS_NO + " " + " PALET NO :" + MBOBINNO + " " + " Gravur Kalite Hata";
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
                    mess.Attachments.Add(new Attachment(fname));
                    smtp.Send(mess);
                }
                komut.CommandText = "INSERT INTO EGEM_GRAVUR_KALITENOKIMAJ (SIPARISNO,BOBINNO,HATA,IMAJ,ADRES) VALUES (@MSIPARIS_NO,@MBOBINNO,@MHATA,@IMAJ,@ADRES)";
                komut.Connection = conn1;
                komut.CommandType = CommandType.Text;
                conn1.Open();
                komut.Parameters.Add("@MSIPARIS_NO", MSIPARIS_NO);
                komut.Parameters.Add("@MBOBINNO", MBOBINNO);
                komut.Parameters.Add("@MHATA", MHATA);
                komut.Parameters.Add("@IMAJ", MIMAJ);
                komut.Parameters.Add("@ADRES", "\\192.168.0.252\\ofisdata\\NOK_IMAJ\\GRAVUR\\" + adi + "_" + MIMAJ);
                komut.ExecuteReader();
                conn1.Close();
                return View();
            }
            catch (Exception)
            {
                return View();
            }
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
                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            fname = testfiles[testfiles.Length - 1];
                        }
                        else
                        {
                            string dosyaadi = DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();
                            string fileName = System.IO.Path.GetFileName(file.FileName);
                            fname = Path.Combine(("//192.168.0.252//ofisdata//NOK_IMAJ//GRAVUR//"), dosyaadi + "_" + fileName);
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


        public ActionResult Sil(int id)
            {
                var baski = db.EGEM_GRAVUR_KALITE.Find(id);
                db.EGEM_GRAVUR_KALITE.Remove(baski);
                db.SaveChanges();
                return RedirectToAction("Gravurduzen");

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

            var customers = from customer in entities.EGEM_GRAVUR_KALITE.Take(10)
                            select customer;

            foreach (var customer in customers)
            {
                dt.Rows.Add(customer.tarih, customer.sipno, customer.urun_kodu, customer.urun_adi, customer.tabaka_hasar);
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