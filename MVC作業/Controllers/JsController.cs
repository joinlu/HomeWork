using MVC作業.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MVC作業.Controllers
{
    public class JsController : Controller
    {
        客戶資料Repository rep客戶資料 = RepositoryHelper.Get客戶資料Repository();

        // GET: Js
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public JsonResult  TestJs()
        {
            var 客戶資料 = rep客戶資料.fn取得所有資料();
            //List<Object> ja = new List<object>();
            List<Dictionary<string, object>> ja = new List<Dictionary<string, object>>();
            foreach (var item in 客戶資料)
            {
                //var jo = new JObject();
                Dictionary<string, object> jo = new Dictionary<string, object>();

                jo.Add("客戶名稱", item.客戶名稱);
                jo.Add("客戶分類", item.客戶分類);
                jo.Add("統一編號", item.統一編號);
                jo.Add("電話", item.電話);
                jo.Add("傳真", item.傳真);
                jo.Add("地址", item.地址);
                jo.Add("Email", item.Email);
                ja.Add(jo);
            }
            return Json(ja);
        }
    }
}