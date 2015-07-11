using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using HW.Models;

namespace HW.Controllers
{
    public class V_ReportController : Controller
    {
        private CustomerEntities db = new CustomerEntities();

        // GET: V_Report
        public ActionResult Index()
        {
            return View(db.V_Report.ToList());
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
