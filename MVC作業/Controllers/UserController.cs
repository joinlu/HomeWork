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
        //private 客戶資料Entities db = new 客戶資料Entities();
        客戶資料Repository rep客戶資料 = RepositoryHelper.Get客戶資料Repository();

        // GET: User
        public ActionResult Index(string namesearch)
        {
            ViewBag.客戶分類 = new SelectList(rep客戶資料.fn分類清單());

            if (namesearch != null && !namesearch.Equals(""))
            {
                var 客戶 = rep客戶資料.fn客戶名稱搜尋(namesearch).ToList();

                if (客戶.Count() > 0)
                {
                    return View(客戶);
                }
                else
                {
                    TempData["回應"] = namesearch;
                    return View(rep客戶資料.fn取得所有資料());
                }
            }
            else
            {
                return View(rep客戶資料.fn取得所有資料());
            }
        }

        // GET: User/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            var 客戶資料 = rep客戶資料.fn單筆搜尋((int)id);
            //客戶資料 客戶資料 = db.客戶資料.Find(id);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        public ActionResult SelectC(string 客戶分類)
        {
            ViewBag.客戶分類 = new SelectList(rep客戶資料.fn分類清單());

            if (客戶分類 != null && !客戶分類.Equals(""))
            {
                if (客戶分類.Equals("全部分類"))
                {
                    return View("Index", rep客戶資料.fn取得所有資料());
                }
                else
                {
                    return View("Index", rep客戶資料.fn客戶篩選資料(客戶分類));
                }
            }
            else
            {
                return View("Index", rep客戶資料.fn取得所有資料());
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
                rep客戶資料.Add(客戶資料);
                rep客戶資料.UnitOfWork.Commit();
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
            var 客戶資料 = rep客戶資料.fn單筆搜尋((int)id);
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
                var 客戶 = rep客戶資料.fn單筆搜尋(客戶資料.Id);

                客戶.傳真 = 客戶資料.傳真;
                客戶.地址 = 客戶資料.地址;
                客戶.客戶分類 = 客戶資料.客戶分類;
                客戶.客戶名稱 = 客戶資料.客戶名稱;
                客戶.客戶聯絡人 = 客戶資料.客戶聯絡人;
                客戶.客戶銀行資訊 = 客戶資料.客戶銀行資訊;
                客戶.統一編號 = 客戶資料.統一編號;
                客戶.電話 = 客戶資料.電話;

                rep客戶資料.UnitOfWork.Commit();

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
            var 客戶資料 = rep客戶資料.fn單筆搜尋((int)id);
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
            var 客戶資料 = rep客戶資料.fn單筆搜尋((int)id);

            客戶資料.IsDeleted = true;

            rep客戶資料.UnitOfWork.Commit();

            return RedirectToAction("Index");
        }

        public ActionResult DSort(string 排序)
        {
            ViewBag.客戶分類 = new SelectList(rep客戶資料.fn分類清單());

            ViewBag.NameSort = 排序 == "Name" ? "Name desc" : "Name";
            ViewBag.UcSort = 排序 == "Uc" ? "Uc desc" : "Uc";
            ViewBag.NumlSort = 排序 == "Num" ? "Num desc" : "Num";
            ViewBag.TelSort = 排序 == "Tel" ? "Tel desc" : "Tel";
            ViewBag.FaxSort = 排序 == "Fax" ? "Fax desc" : "Fax";
            ViewBag.AddSort = 排序 == "Add" ? "Add desc" : "Add";
            ViewBag.EmailSort = 排序 == "Email" ? "Email desc" : "Email";

            return View("Index", rep客戶資料.fn重新排序(排序).ToList());
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

            var customers = rep客戶資料.fn取得所有資料();

            foreach (var customer in customers)
            {
                dt.Rows.Add(customer.客戶名稱, customer.客戶分類, customer.統一編號, customer.電話, customer.傳真, customer.地址, customer.Email);
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
                //db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
