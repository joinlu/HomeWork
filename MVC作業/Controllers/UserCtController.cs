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
    public class UserCtController : Controller
    {
        private 客戶資料Entities db = new 客戶資料Entities();

        // GET: UserCt
        public ActionResult Index(string namesearch)
        {
            List<String> item職稱 = new List<string>();
            item職稱.Add("全部職稱");
            var 職稱清單 = (from data in db.客戶聯絡人
                        where data.IsDeleted == false
                        select data.職稱);
            foreach (var item in 職稱清單.Distinct().ToList())
            {
                item職稱.Add(item.ToString());
            }
            ViewBag.職稱 = new SelectList(item職稱);

            if (namesearch != null && !namesearch.Equals(""))
            {
                var 客戶聯絡人資料 = db.客戶聯絡人.Where(x => x.姓名.Equals(namesearch) && x.IsDeleted == false).ToList();

                if (客戶聯絡人資料.Count > 0)
                {
                    return View(客戶聯絡人資料);
                }
                else
                {
                    TempData["回應"] = namesearch;
                    var 客戶聯絡人 = db.客戶聯絡人.Where(x => x.IsDeleted == false).Include(客 => 客.客戶資料);
                    return View(客戶聯絡人.ToList());
                }
            }
            else
            {
                var 客戶聯絡人 = db.客戶聯絡人.Include(客 => 客.客戶資料);
                return View(客戶聯絡人.ToList());
            }
        }

        public ActionResult SelectC(string 職稱)
        {
            List<String> item職稱 = new List<string>();
            item職稱.Add("全部職稱");
            var 職稱清單 = (from data in db.客戶聯絡人
                        where data.IsDeleted == false
                        select data.職稱);
            foreach (var item in 職稱清單.Distinct().ToList())
            {
                item職稱.Add(item.ToString());
            }
            ViewBag.職稱 = new SelectList(item職稱);

            if (職稱 != null && !職稱.Equals(""))
            {
                if (職稱.Equals("全部職稱"))
                {
                    var 客戶聯絡人資料 = db.客戶聯絡人.Where(x => x.IsDeleted == false).ToList();
                    return View("Index",客戶聯絡人資料);
                }
                else
                {
                    var 客戶聯絡人資料 = db.客戶聯絡人.Where(x => x.職稱.Equals(職稱) && x.IsDeleted == false).ToList();
                    return View("Index", 客戶聯絡人資料);
                }
            }
            else
            {
                var 客戶聯絡人資料 = db.客戶聯絡人.Where(x => x.IsDeleted == false).ToList();
                return View("Index", 客戶聯絡人資料);
            }
        }

        public ActionResult DSort(string 排序)
        {
            List<String> item職稱 = new List<string>();
            item職稱.Add("全部職稱");
            var 職稱清單 = (from data in db.客戶聯絡人
                        where data.IsDeleted == false
                        select data.職稱);
            foreach (var item in 職稱清單.Distinct().ToList())
            {
                item職稱.Add(item.ToString());
            }
            ViewBag.職稱 = new SelectList(item職稱);


            ViewBag.NameSort = String.IsNullOrEmpty(排序) ? "Name desc" : "";
            ViewBag.EmailSort = 排序 == "Email" ? "Email desc" : "Email";
            ViewBag.PhoneSort = 排序 == "Phone" ? "Phone desc" : "Phone";
            ViewBag.TelSort = 排序 == "Tel" ? "Tel desc" : "Tel";
            ViewBag.UserSort = 排序 == "User" ? "User desc" : "User";
            ViewBag.JobSort = 排序 == "Job" ? "Job desc" : "Job";

            var 客戶聯絡人 = db.客戶聯絡人.Where(x => x.IsDeleted == false).Include(客 => 客.客戶資料);
            switch (排序)
            {
                case "Name desc":
                    客戶聯絡人 = 客戶聯絡人.OrderByDescending(s => s.姓名);
                    break;
                case "Name":
                    客戶聯絡人 = 客戶聯絡人.OrderBy(s => s.姓名);
                    break;
                case "Email desc":
                    客戶聯絡人 = 客戶聯絡人.OrderByDescending(s => s.Email);
                    break;
                case "Email":
                    客戶聯絡人 = 客戶聯絡人.OrderBy(s => s.Email);
                    break;
                case "Phone desc":
                    客戶聯絡人 = 客戶聯絡人.OrderByDescending(s => s.手機);
                    break;
                case "Phone":
                    客戶聯絡人 = 客戶聯絡人.OrderBy(s => s.手機);
                    break;
                case "Tel desc":
                    客戶聯絡人 = 客戶聯絡人.OrderByDescending(s => s.電話);
                    break;
                case "Tel":
                    客戶聯絡人 = 客戶聯絡人.OrderBy(s => s.電話);
                    break;
                case "User desc":
                    客戶聯絡人 = 客戶聯絡人.OrderByDescending(s => s.客戶Id);
                    break;
                case "User":
                    客戶聯絡人 = 客戶聯絡人.OrderBy(s => s.客戶Id);
                    break;
                case "Job desc":
                    客戶聯絡人 = 客戶聯絡人.OrderByDescending(s => s.職稱);
                    break;
                case "Job":
                    客戶聯絡人 = 客戶聯絡人.OrderBy(s => s.職稱);
                    break;

                default:
                    break;
            }


            return View("Index",客戶聯絡人.ToList());

        }

        // GET: UserCt/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶聯絡人 客戶聯絡人 = db.客戶聯絡人.Find(id);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            return View(客戶聯絡人);
        }

        // GET: UserCt/Create
        public ActionResult Create()
        {
            ViewBag.客戶Id = new SelectList(db.客戶資料, "Id", "客戶名稱");
            return View();
        }

        // POST: UserCt/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 https://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶Id,職稱,姓名,Email,手機,電話")] 客戶聯絡人 客戶聯絡人)
        {
            if (ModelState.IsValid)
            {
                var 客戶聯絡人資料 = db.客戶聯絡人.Where(x => x.Email.Equals(客戶聯絡人.Email)).FirstOrDefault();
                if (客戶聯絡人資料 != null)
                {
                    TempData["回應"] = 客戶聯絡人;
                }
                else
                {
                    db.客戶聯絡人.Add(客戶聯絡人);
                    db.SaveChanges();
                    return RedirectToAction("Index");
                }
            }

            ViewBag.客戶Id = new SelectList(db.客戶資料, "Id", "客戶名稱", 客戶聯絡人.客戶Id);
            return View(客戶聯絡人);
        }

        // GET: UserCt/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶聯絡人 客戶聯絡人 = db.客戶聯絡人.Find(id);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            ViewBag.客戶Id = new SelectList(db.客戶資料, "Id", "客戶名稱", 客戶聯絡人.客戶Id);
            return View(客戶聯絡人);
        }

        // POST: UserCt/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 https://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶Id,職稱,姓名,Email,手機,電話")] 客戶聯絡人 客戶聯絡人)
        {
            if (ModelState.IsValid)
            {
                //var 客戶聯絡人資料 = db.客戶聯絡人.Where(x => x.Email.Equals(客戶聯絡人.Email) && x.IsDeleted==false).FirstOrDefault();
                //if (客戶聯絡人資料 != null)
                //{
                //    TempData["回應"] = 客戶聯絡人;
                //}
                //else
                //{
                db.Entry(客戶聯絡人).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
                //}
            }

            ViewBag.客戶Id = new SelectList(db.客戶資料, "Id", "客戶名稱", 客戶聯絡人.客戶Id);
            return View(客戶聯絡人);
        }

        // GET: UserCt/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶聯絡人 客戶聯絡人 = db.客戶聯絡人.Find(id);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            return View(客戶聯絡人);
        }

        // POST: UserCt/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            客戶聯絡人 客戶聯絡人 = db.客戶聯絡人.Find(id);
            //db.客戶聯絡人.Remove(客戶聯絡人);
            客戶聯絡人.IsDeleted = true;
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        public FileResult Export()
        {
            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[5] { new DataColumn("職稱"),
                                                    new DataColumn("姓名"),
                                                    new DataColumn("Email"),
                                                    new DataColumn("手機"),
                                                    new DataColumn("電話")});

            var customers = from customer in db.客戶聯絡人.Where(x => x.IsDeleted == false)
                            select customer;

            foreach (var customer in customers)
            {
                dt.Rows.Add(customer.職稱, customer.姓名, customer.Email, customer.手機, customer.電話);
            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                wb.Worksheet(1).Tables.FirstOrDefault().ShowAutoFilter = false;

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "客戶聯絡資料.xlsx");
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
