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
    public class EsnekLaminasyonController : Controller
    {
        EGEM2021Entities1 db = new EGEM2021Entities1();
        SqlConnection conn1 = new SqlConnection("Data Source=192.168.0.251;Initial Catalog=EGEM2022;Persist Security Info=True;User ID=egem;Password=123456");
        public static string adi;
        public static string fname;
        SqlCommand komut = new SqlCommand();
        public static string kod;


        public ActionResult Laminasyonlist(int sayfa = 1)
        {
            var degerler = db.EGEM_ESNEK_LAMINASYON_KALITE.ToList().ToPagedList(sayfa, 20);
            return View(degerler);
        }

        public ActionResult Laminasyonduzen(string Search_Data, string Filter_Value, int? Page_No)
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
            var students = from stu in db.EGEM_ESNEK_LAMINASYON_KALITE select stu;
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
        public ActionResult Laminasyonekle()
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
        public ActionResult Laminasyonekle(EGEM_ESNEK_LAMINASYON_KALITE p1)
        {
            List<SelectListItem> degerler1 = (from i in db.EGEM_KALITE_PERSONEL.ToList()
                                              select new SelectListItem
                                              {
                                                  Text = i.ADSOYAD
                                              }).ToList();
            ViewBag.KALITE_ONAY = degerler1;

            db.EGEM_ESNEK_LAMINASYON_KALITE.Add(p1);
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
                var subject = "SIPARIS NO : " + MSIPARIS_NO + " " + " PALET NO :" + MBOBINNO + " " + " Esnek Laminasyon Kalite Hata";
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
                komut.CommandText = "INSERT INTO EGEM_ESNEK_LAMINASYON_KALITENOKIMAJ (SIPARISNO,BOBINNO,HATA,IMAJ,ADRES) VALUES (@MSIPARIS_NO,@MBOBINNO,@MHATA,@IMAJ,@ADRES)";
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
                    komut.Parameters.Add("@ADRES", adi + "_" + MIMAJ);
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
            var laminasyon = db.EGEM_ESNEK_LAMINASYON_KALITE.Find(id);
            db.EGEM_ESNEK_LAMINASYON_KALITE.Remove(laminasyon);
            db.SaveChanges();
            return RedirectToAction("Laminasyonduzen");

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
                            fname = Path.Combine(("~/Content/NOK_IMAJ/ESNEKLAMINASYON/"), dosyaadi + "_" + fileName);
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

        public ActionResult LaminasyonNOK(String Search_Data, String Filter_Value, int? Page_No)
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

            var students = from stu in db.EGEM_ESNEK_LAMINASYON_KALITENOKIMAJ select stu;
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
        public FileResult Export()
        {
            EGEM2021Entities1 entities = new EGEM2021Entities1();
            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[17]
            {
                new DataColumn("URUN_KODU"),
                new DataColumn("URUN_ADI"),
                new DataColumn("SIPARIS_NO"),
                new DataColumn("ZAMAN"),
                new DataColumn("BOBIN NO"),
                new DataColumn("BOBIN ENI"),
                new DataColumn("FOTOSEL ARALIGI"),
                new DataColumn("AL_FOLYO_YONU"),
                new DataColumn("TUTKAL GRAMAJI"),
                new DataColumn("BIRIM AGIRLIK"),
                new DataColumn("YAPISMA KONTROLU"),
                new DataColumn("GORSEL KONTROL"),
                new DataColumn("BENCIK KONTROLU"),
                new DataColumn("ISIL YAPISMA KONTROLU"),
                new DataColumn("LAMINASYON KUVVETI"),
                new DataColumn("COF TESTI"),
                new DataColumn("KALITE ONAY"),
            });
            var customers = from stu in db.EGEM_ESNEK_LAMINASYON_KALITE select stu;
            customers = customers.Where(stu => stu.ZAMAN.ToString().Contains(kod.ToString()));
            foreach (var customer in customers.ToList())
            {
                dt.Rows.Add(
                    customer.URUN_KODU,
                    customer.URUN_ADI,
                    customer.SIPARIS_NO,
                    customer.ZAMAN,
                    customer.BOBIN_NO,
                    customer.BOBIN_ENI,
                    customer.FOTOSEL_ARALIGI,
                    customer.AL_FOLYO_YONU,
                    customer.TUTKAL_GRAMAJ,
                    customer.BIRIM_AGIRLIK,
                    customer.YAPISMA_KONTROL,
                    customer.GORSEL_KONTROL,
                    customer.BENCIK_KONTROL,
                    customer.ISIL_YAPISMA_KONT,
                    customer.LAMINASYON_KUVVETI,
                    customer.COF_TEST,
                    customer.KALITE_ONAY
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
