using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using MVC作業.Models;
using ClosedXML.Excel;
using System.IO;

namespace MVC作業.Controllers
{
    public class UserController : Controller
    {
        private 客戶資料Entities db = new 客戶資料Entities();

        // GET: User
        public ActionResult Index(string namesearch)
        {
            List<String> item客戶分類 = new List<string>();
            item客戶分類.Add("全部分類");
            var 分類清單 = (from data in db.客戶資料
                        where data.IsDeleted == false
                        select data.客戶分類).Distinct().ToList();
            foreach (var item in 分類清單)
            {
                if (item != null)
                {
                    item客戶分類.Add(item.ToString());
                }
            }
            ViewBag.客戶分類 = new SelectList(item客戶分類);


            if (namesearch != null && !namesearch.Equals(""))
            {
                var 客戶 = db.客戶資料.Where(x => x.客戶名稱.Equals(namesearch) && x.IsDeleted == false).ToList();

                if (客戶.Count > 0)
                {
                    return View(客戶);
                }
                else
                {
                    TempData["回應"] = namesearch;
                    return View(db.客戶資料.Where(x => x.IsDeleted == false).ToList());
                }
            }
            else
            {
                return View(db.客戶資料.Where(x => x.IsDeleted == false).ToList());
            }

        }

        // GET: User/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = db.客戶資料.Find(id);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        public ActionResult SelectC(string 客戶分類)
        {
            List<String> item客戶分類 = new List<string>();
            item客戶分類.Add("全部分類");
            var 分類清單 = (from data in db.客戶資料
                        where data.IsDeleted == false
                        select data.客戶分類).Distinct().ToList();
            foreach (var item in 分類清單)
            {
                if (item != null)
                {
                    item客戶分類.Add(item.ToString());
                }
            }
            ViewBag.客戶分類 = new SelectList(item客戶分類);

            if (客戶分類 != null && !客戶分類.Equals(""))
            {
                if (客戶分類.Equals("全部分類"))
                {
                    var 客戶篩選資料 = db.客戶資料.Where(x => x.IsDeleted == false).ToList();
                    return View("Index", 客戶篩選資料);
                }
                else
                {
                    var 客戶篩選資料 = db.客戶資料.Where(x => x.客戶分類.Equals(客戶分類) && x.IsDeleted == false).ToList();
                    return View("Index", 客戶篩選資料);
                }
            }
            else
            {
                var 客戶篩選資料 = db.客戶資料.Where(x => x.IsDeleted == false).ToList();
                return View("Index", 客戶篩選資料);
            }
        }

        // GET: User/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: User/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 https://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶名稱,客戶分類,統一編號,電話,傳真,地址,Email")] 客戶資料 客戶資料)
        {
            if (ModelState.IsValid)
            {
                db.客戶資料.Add(客戶資料);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(客戶資料);
        }

        // GET: User/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = db.客戶資料.Find(id);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        // POST: User/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 https://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶名稱,客戶分類,統一編號,電話,傳真,地址,Email")] 客戶資料 客戶資料)
        {
            if (ModelState.IsValid)
            {
                db.Entry(客戶資料).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(客戶資料);
        }

        // GET: User/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = db.客戶資料.Find(id);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        // POST: User/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            客戶資料 客戶資料 = db.客戶資料.Find(id);
            //db.客戶資料.Remove(客戶資料);
            客戶資料.IsDeleted = true;
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        public ActionResult DSort(string 排序)
        {
            List<String> item客戶分類 = new List<string>();
            item客戶分類.Add("全部分類");
            var 分類清單 = (from data in db.客戶資料
                        where data.IsDeleted == false
                        select data.客戶分類);
            foreach (var item in 分類清單.Distinct().ToList())
            {
                item客戶分類.Add(item.ToString());
            }
            ViewBag.職稱 = new SelectList(item客戶分類);


            ViewBag.NameSort = 排序 == "Name" ? "Name desc" : "Name";
            ViewBag.UcSort = 排序 == "Uc" ? "Uc desc" : "Uc";
            ViewBag.NumlSort = 排序 == "Num" ? "Num desc" : "Num";
            ViewBag.TelSort = 排序 == "Tel" ? "Tel desc" : "Tel";
            ViewBag.FaxSort = 排序 == "Fax" ? "Fax desc" : "Fax";
            ViewBag.AddSort = 排序 == "Add" ? "Add desc" : "Add";
            ViewBag.EmailSort = 排序 == "Email" ? "Email desc" : "Email";

            var 客戶資料 = db.客戶資料.Where(x => x.IsDeleted == false);
            switch (排序)
            {
                case "Name desc":
                    客戶資料 = 客戶資料.OrderByDescending(s => s.客戶名稱);
                    break;
                case "Name":
                    客戶資料 = 客戶資料.OrderBy(s => s.客戶名稱);
                    break;
                case "Uc desc":
                    客戶資料 = 客戶資料.OrderByDescending(s => s.客戶分類);
                    break;
                case "Uc":
                    客戶資料 = 客戶資料.OrderBy(s => s.客戶分類);
                    break;
                case "Num desc":
                    客戶資料 = 客戶資料.OrderByDescending(s => s.統一編號);
                    break;
                case "Num":
                    客戶資料 = 客戶資料.OrderBy(s => s.統一編號);
                    break;
                case "Tel desc":
                    客戶資料 = 客戶資料.OrderByDescending(s => s.電話);
                    break;
                case "Tel":
                    客戶資料 = 客戶資料.OrderBy(s => s.電話);
                    break;
                case "Fax desc":
                    客戶資料 = 客戶資料.OrderByDescending(s => s.傳真);
                    break;
                case "Fax":
                    客戶資料 = 客戶資料.OrderBy(s => s.傳真);
                    break;
                case "Add desc":
                    客戶資料 = 客戶資料.OrderByDescending(s => s.地址);
                    break;
                case "Add":
                    客戶資料 = 客戶資料.OrderBy(s => s.地址);
                    break;
                case "Email desc":
                    客戶資料 = 客戶資料.OrderByDescending(s => s.Email);
                    break;
                case "Email":
                    客戶資料 = 客戶資料.OrderBy(s => s.Email);
                    break;

                default:
                    break;
            }

            return View("Index", 客戶資料.ToList());
        }


        //[HttpPost]
        public FileResult Export()
        {
            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[7] { new DataColumn("客戶名稱"),
                                                    new DataColumn("客戶分類"),
                                                    new DataColumn("統一編號"),
                                                    new DataColumn("電話"),
                                                    new DataColumn("傳真"),
                                                    new DataColumn("地址"),
                                                    new DataColumn("Email") });

            var customers = from customer in db.客戶資料.Where(x => x.IsDeleted == false)
                            select customer;

            foreach (var customer in customers)
            {
                dt.Rows.Add(customer.客戶名稱,customer.客戶分類, customer.統一編號, customer.電話, customer.傳真, customer.地址, customer.Email);
            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                wb.Worksheet(1).Tables.FirstOrDefault().ShowAutoFilter = false;

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "客戶資料.xlsx");
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
