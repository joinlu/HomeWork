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
    public class UserBankController : Controller
    {
        //private 客戶資料Entities db = new 客戶資料Entities();
        客戶銀行資訊Repository rep客戶銀行資訊 = RepositoryHelper.Get客戶銀行資訊Repository();
        客戶資料Repository rep客戶資料 = RepositoryHelper.Get客戶資料Repository();

        // GET: UserBank
        public ActionResult Index(string namesearch)
        {
            if (namesearch != null && !namesearch.Equals(""))
            {
                var 銀行資訊 = rep客戶銀行資訊.fn銀行名稱搜尋(namesearch).ToList();

                if (銀行資訊.Count > 0)
                {
                    return View(銀行資訊);
                }
                else
                {
                    TempData["回應"] = namesearch;
                    var 客戶銀行資訊 = rep客戶銀行資訊.fn取得所有資料().Include(客 => 客.客戶資料);
                    return View(客戶銀行資訊);
                }

            }
            else
            {
                var 客戶銀行資訊 = rep客戶銀行資訊.fn取得所有資料().Include(客 => 客.客戶資料);
                return View(客戶銀行資訊);
            }
        }

        // GET: UserBank/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            var 客戶銀行資訊 = rep客戶銀行資訊.fn單筆搜尋((int)id);
            if (客戶銀行資訊 == null)
            {
                return HttpNotFound();
            }
            return View(客戶銀行資訊);
        }

        // GET: UserBank/Create
        public ActionResult Create()
        {
            ViewBag.客戶Id = new SelectList(rep客戶資料.fn取得所有資料(), "Id", "客戶名稱");
            return View();
        }

        // POST: UserBank/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 https://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶Id,銀行名稱,銀行代碼,分行代碼,帳戶名稱,帳戶號碼")] 客戶銀行資訊 客戶銀行資訊)
        {
            if (ModelState.IsValid)
            {
                rep客戶銀行資訊.Add(客戶銀行資訊);
                rep客戶銀行資訊.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            ViewBag.客戶Id = new SelectList(rep客戶資料.fn取得所有資料(), "Id", "客戶名稱");
            return View(客戶銀行資訊);
        }

        // GET: UserBank/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            //var 客戶銀行資訊 = rep客戶銀行資訊.fn單筆搜尋編輯用((int)id).Include(客 => 客.客戶資料).ToList();
            var 客戶銀行資訊 = rep客戶銀行資訊.fn單筆搜尋((int)id);
            
            if (客戶銀行資訊 == null)
            {
                return HttpNotFound();
            }
            SelectList selectList = new SelectList(rep客戶資料.fn取得所有資料(), "Id", "客戶名稱", 客戶銀行資訊.客戶Id);
            //SelectList 
            ViewBag.客戶Id = selectList;
            return View(客戶銀行資訊);
        }

        // POST: UserBank/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 https://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶Id,銀行名稱,銀行代碼,分行代碼,帳戶名稱,帳戶號碼")] 客戶銀行資訊 客戶銀行資訊)
        {
            if (ModelState.IsValid)
            {
                var 銀行 = rep客戶銀行資訊.fn單筆搜尋(客戶銀行資訊.Id);

                銀行.銀行名稱 = 客戶銀行資訊.銀行名稱;
                銀行.銀行代碼 = 客戶銀行資訊.銀行代碼;
                銀行.分行代碼 = 客戶銀行資訊.分行代碼;
                銀行.帳戶名稱 = 客戶銀行資訊.帳戶名稱;
                銀行.帳戶號碼 = 客戶銀行資訊.帳戶號碼;
                銀行.客戶Id = 客戶銀行資訊.客戶Id;

                rep客戶銀行資訊.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }
            ViewBag.客戶Id = new SelectList(rep客戶資料.fn取得所有資料(), "Id", "客戶名稱");
            return View(客戶銀行資訊);
        }

        // GET: UserBank/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            var 客戶銀行資訊 = rep客戶銀行資訊.fn單筆搜尋((int)id);
            if (客戶銀行資訊 == null)
            {
                return HttpNotFound();
            }
            return View(客戶銀行資訊);
        }

        // POST: UserBank/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            var 客戶銀行資訊 = rep客戶銀行資訊.fn單筆搜尋((int)id);
            客戶銀行資訊.IsDeleted = true;
            rep客戶銀行資訊.UnitOfWork.Commit();
            return RedirectToAction("Index");
        }

        public FileResult Export()
        {
            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[5] { new DataColumn("銀行名稱"),
                                                    new DataColumn("銀行代碼"),
                                                    new DataColumn("分行代碼"),
                                                    new DataColumn("帳戶名稱"),
                                                    new DataColumn("帳戶號碼")});

            var customers = rep客戶銀行資訊.fn取得所有資料();

            foreach (var customer in customers)
            {
                dt.Rows.Add(customer.銀行名稱, customer.銀行代碼, customer.分行代碼, customer.帳戶名稱, customer.帳戶號碼);
            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                wb.Worksheet(1).Tables.FirstOrDefault().ShowAutoFilter = false;

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "客戶銀行資料.xlsx");
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
