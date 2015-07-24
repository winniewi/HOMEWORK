using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using HW.Models;
using NPOI.HSSF.UserModel;
using System.IO;

namespace HW.Controllers
{
    public class 客戶資料Controller : Controller
    {
        private CustomerEntities db = new CustomerEntities();

        // GET: 客戶資料
        public ActionResult Index()
        {
            return View(db.客戶資料.Where(p => p.是否已刪除 == false).ToList());
        }

        // GET: 客戶資料/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = db.客戶資料.Find(id);
            if (客戶資料 == null && 客戶資料.是否已刪除 == false)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        // GET: 客戶資料/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: 客戶資料/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶名稱,統一編號,電話,傳真,地址,Email")] 客戶資料 客戶資料)
        {
            if (ModelState.IsValid)
            {
                if (this.是否重複(客戶資料.Id, 客戶資料.Email) == true)
                {
                    ModelState.AddModelError("Email", new Exception("Email重複"));
                }
                else
                {
                    db.客戶資料.Add(客戶資料);
                    db.SaveChanges();
                    return RedirectToAction("Index");
                }
            }

            return View(客戶資料);
        }
        private bool 是否重複(int id, string email)
        {
            return this.db.客戶資料.Any(p => p.Email == email && p.Id != id && p.是否已刪除 == false);
        }

        // GET: 客戶資料/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = db.客戶資料.Find(id);
            if (客戶資料 == null && 客戶資料.是否已刪除 == false)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        // POST: 客戶資料/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶名稱,統一編號,電話,傳真,地址,Email")] 客戶資料 客戶資料)
        {
            if (ModelState.IsValid)
            {
                if (this.是否重複(客戶資料.Id, 客戶資料.Email) == true)
                {
                    ModelState.AddModelError("Email", new Exception("Email重複"));
                }
                else
                {
                    db.Entry(客戶資料).State = EntityState.Modified;
                    db.SaveChanges();
                    return RedirectToAction("Index");
                }
            }
            return View(客戶資料);
        }

        // GET: 客戶資料/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = db.客戶資料.Find(id);
            if (客戶資料 == null && 客戶資料.是否已刪除 == false)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        public ActionResult Export()
        {
            var data = db.客戶資料.Where(p => p.是否已刪除 == false).ToList();
            byte[] _bytes = null;
            var workbook = new HSSFWorkbook();
            var sheet = workbook.CreateSheet("Sheet1");
            var row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("客戶名稱");
            row.CreateCell(1).SetCellValue("統一編號");
            row.CreateCell(2).SetCellValue("電話");
            row.CreateCell(3).SetCellValue("傳真");
            row.CreateCell(4).SetCellValue("地址");
            row.CreateCell(5).SetCellValue("Email");
            for (int i = 0; i < data.Count; i++)
            {
                row = sheet.CreateRow(i + 1);
                row.CreateCell(0).SetCellValue(data[i].客戶名稱);
                row.CreateCell(1).SetCellValue(data[i].統一編號);
                row.CreateCell(2).SetCellValue(data[i].電話);
                row.CreateCell(3).SetCellValue(data[i].傳真);
                row.CreateCell(4).SetCellValue(data[i].地址);
                row.CreateCell(5).SetCellValue(data[i].Email);
            }
            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                stream.Position = 0;
                stream.Flush();
                _bytes = stream.GetBuffer();
            }
            return File(_bytes, "application/vnd.ms-excel");
        }

        [HttpPost]
        public ActionResult BatchDelete(int[] chkDelete)
        {
            foreach (var id in chkDelete)
            {
                var q = db.客戶資料.Find(id);
                q.是否已刪除 = true;
            }
            if (db.ChangeTracker.HasChanges() == true)
            {
                db.SaveChanges();
            }
            return Redirect("Index");
        }

        // POST: 客戶資料/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            客戶資料 客戶資料 = db.客戶資料.Find(id);
            客戶資料.是否已刪除 = true;
            db.SaveChanges();
            return RedirectToAction("Index");
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
