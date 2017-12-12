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
        //private 客戶資料Entities db = new 客戶資料Entities();

        客戶聯絡人Repository rep客戶聯絡人 = RepositoryHelper.Get客戶聯絡人Repository();
        客戶資料Repository rep客戶資料 = RepositoryHelper.Get客戶資料Repository();

        // GET: UserCt
        public ActionResult Index(string namesearch)
        {
            ViewBag.職稱 = new SelectList(rep客戶聯絡人.fn職稱清單());

            if (namesearch != null && !namesearch.Equals(""))
            {
                var 客戶聯絡人資料 =rep客戶聯絡人.fn客戶名稱搜尋(namesearch).ToList();

                if (客戶聯絡人資料.Count > 0)
                {
                    return View(客戶聯絡人資料);
                }
                else
                {
                    TempData["回應"] = namesearch;
                    return View(rep客戶聯絡人.fn取得所有資料());
                }
            }
            else
            {
                return View(rep客戶聯絡人.fn取得所有資料());
            }
        }

        public ActionResult SelectC(string 職稱)
        {
            ViewBag.職稱 = new SelectList(rep客戶聯絡人.fn職稱清單());

            if (職稱 != null && !職稱.Equals(""))
            {
                if (職稱.Equals("全部職稱"))
                {
                    return View("Index", rep客戶聯絡人.fn取得所有資料());
                }
                else
                {
                   return View("Index",rep客戶聯絡人.fn客戶篩選資料(職稱));
                }
            }
            else
            {
                return View("Index", rep客戶聯絡人.fn取得所有資料());
            }
        }

        public ActionResult DSort(string 排序)
        {
            ViewBag.職稱 = new SelectList(rep客戶聯絡人.fn職稱清單());

            ViewBag.NameSort = String.IsNullOrEmpty(排序) ? "Name desc" : "";
            ViewBag.EmailSort = 排序 == "Email" ? "Email desc" : "Email";
            ViewBag.PhoneSort = 排序 == "Phone" ? "Phone desc" : "Phone";
            ViewBag.TelSort = 排序 == "Tel" ? "Tel desc" : "Tel";
            ViewBag.UserSort = 排序 == "User" ? "User desc" : "User";
            ViewBag.JobSort = 排序 == "Job" ? "Job desc" : "Job";

            return View("Index", rep客戶聯絡人.fn重新排序(排序).ToList());
        }

        // GET: UserCt/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            //客戶聯絡人 客戶聯絡人 = db.客戶聯絡人.Find(id);
            var 客戶聯絡人 = rep客戶聯絡人.fn單筆搜尋((int)id);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            return View(客戶聯絡人);
        }

        // GET: UserCt/Create
        public ActionResult Create()
        {
            
            ViewBag.客戶Id = new SelectList(rep客戶資料.fn取得所有資料(), "Id", "客戶名稱");
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
                var 客戶聯絡人資料 = rep客戶聯絡人.fnEail是否已存在(客戶聯絡人.Email);
                if (客戶聯絡人資料 != null)
                {
                    TempData["回應"] = 客戶聯絡人;
                }
                else
                {
                    rep客戶聯絡人.Add(客戶聯絡人);
                    rep客戶聯絡人.UnitOfWork.Commit();
                    return RedirectToAction("Index");
                }
            }

            ViewBag.客戶Id = new SelectList(rep客戶資料.fn取得所有資料(), "Id", "客戶名稱");
            return View(客戶聯絡人);
        }

        // GET: UserCt/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            var 客戶聯絡人 = rep客戶聯絡人.fn單筆搜尋((int)id);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            ViewBag.客戶Id = new SelectList(rep客戶資料.fn取得所有資料(), "Id", "客戶名稱", 客戶聯絡人.客戶Id);
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
                var 客戶 = rep客戶聯絡人.fn單筆搜尋(客戶聯絡人.Id);

                客戶.職稱 = 客戶聯絡人.職稱;
                客戶.姓名 = 客戶聯絡人.姓名;
                客戶.Email = 客戶聯絡人.Email;
                客戶.手機 = 客戶聯絡人.手機;
                客戶.電話 = 客戶聯絡人.電話;
                客戶.客戶Id = 客戶聯絡人.客戶Id;

                rep客戶聯絡人.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            ViewBag.客戶Id = new SelectList(rep客戶資料.fn取得所有資料(), "Id", "客戶名稱");
            return View(客戶聯絡人);
        }

        // GET: UserCt/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            var 客戶聯絡人 = rep客戶聯絡人.fn單筆搜尋((int)id);
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
            var 客戶聯絡人 = rep客戶聯絡人.fn單筆搜尋((int)id);

            客戶聯絡人.IsDeleted = true;

            rep客戶聯絡人.UnitOfWork.Commit();

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

            var customers = rep客戶聯絡人.fn取得所有資料();

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
                //db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
