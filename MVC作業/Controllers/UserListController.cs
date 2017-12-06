using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using MVC作業.Models;

namespace MVC作業.Controllers
{
    public class UserListController : Controller
    {
        private 客戶資料Entities db = new 客戶資料Entities();

        // GET: UserList
        public ActionResult Index()
        {
            return View(db.UserListView.ToList());
        }

    }
}
