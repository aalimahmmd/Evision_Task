using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using e_vision.Models;
using System.Drawing;
using System.IO;
using System.Data.Entity.Infrastructure;
using OfficeOpenXml;

namespace e_vision.Controllers
{
    public class ProductsController : Controller
    {
        private ProductContext db = new ProductContext();

        // GET: Products
        public ActionResult Index(string searchBy, string search)
        {
            List<ProductViewModel> productlist = db.Products.Select(x => new ProductViewModel
            {
                Id = x.Id,
                Name = x.Name,
                Price = x.Price,
                Photo = x.Photo,
                LastUpdated = x.LastUpdated
            }).ToList();

            if (searchBy == "Price")
            {
                int code = Convert.ToInt32(search);
                return View(db.Products.Where(x => x.Price == code || search == null).ToList());
            }
            else
            {
                return View(db.Products.Where(x => x.Name.StartsWith(search) || search == null).ToList());
            }
        }

        public void ExportToExcel()
        {
            List<ProductViewModel> productlist = db.Products.Select(x => new ProductViewModel
            {
                Id = x.Id,
                Name = x.Name,
                Price = x.Price,
                Photo = x.Photo,
                LastUpdated = x.LastUpdated
            }).ToList();

            ExcelPackage epkg = new ExcelPackage();
            ExcelWorksheet ws = epkg.Workbook.Worksheets.Add("Report");

            ws.Cells["A1"].Value = "Communication";
            ws.Cells["B1"].Value = "Com1";

            ws.Cells["A2"].Value = "Report";
            ws.Cells["B2"].Value = "Report1";

            ws.Cells["A3"].Value = "Date";
            ws.Cells["B3"].Value = string.Format("{0:dd MMMM yyyy} at {0:H: mm tt}",DateTimeOffset.Now);

            ws.Cells["A6"].Value = "Id";
            ws.Cells["B6"].Value = "Name";
            ws.Cells["C6"].Value = "Photo";
            ws.Cells["D6"].Value = "Price";
            ws.Cells["E6"].Value = "LastUpdated";

            int rowStart = 7;
            foreach (var itm in productlist)
            {
                ws.Cells[string.Format("A{0}", rowStart)].Value = itm.Id;
                ws.Cells[string.Format("B{0}", rowStart)].Value = itm.Name;
                ws.Cells[string.Format("C{0}", rowStart)].Value = itm.Photo;
                ws.Cells[string.Format("D{0}", rowStart)].Value = itm.Price;
                ws.Cells[string.Format("E{0}", rowStart)].Value = itm.LastUpdated;
                rowStart++;
            }

            ws.Cells["A:AZ"].AutoFitColumns();
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment: filename=" + "ExcelReport.xlsx");
            Response.BinaryWrite(epkg.GetAsByteArray());
            Response.End();
        }

        // GET: Products/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Product product = db.Products.Find(id);
            if (product == null)
            {
                return HttpNotFound();
            }
            return View(product);
        }

        // GET: Products/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Products/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,Name,Photo,Price,LastUpdated")] Product product, HttpPostedFileBase img)
        {
            if (ModelState.IsValid)
            {
                string fileName = Path.GetFileName(img.FileName);
                var path = Path.Combine(Server.MapPath("~/Photos/"), fileName);
                img.SaveAs(path);
                string trans2 = fileName;
                product.Photo = trans2;
                db.Products.Add(product);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(product);
        }

        // GET: Products/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Product product = db.Products.Find(id);
            if (product == null)
            {
                return HttpNotFound();
            }
            return View(product);
        }

        // POST: Products/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,Name,Photo,Price,LastUpdated")] Product product, HttpPostedFileBase img)
        {
            if (ModelState.IsValid)
            {
                string fileName = Path.GetFileName(img.FileName);
                var path = Path.Combine(Server.MapPath("~/Photos/"), fileName);
                img.SaveAs(path);
                string trans2 = fileName;
                product.Photo = trans2;
                db.Entry(product).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(product);
        }

        // GET: Products/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Product product = db.Products.Find(id);
            if (product == null)
            {
                return HttpNotFound();
            }
            return View(product);
        }

        // POST: Products/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Product product = db.Products.Find(id);
            db.Products.Remove(product);
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

       /* public ActionResult SaveData(Product product)
        {
            if(product.Name != null && product.Price != null && product.Photo != null && product.LastUpdated != null)
            {
                string fileName = Path.GetFileNameWithoutExtension(product.ImageFile.FileName);
                string extension = Path.GetExtension(product.ImageFile.FileName);
                fileName = fileName + DateTime.Now.ToString("yymmssfff") + extension;
                product.ImageFile.SaveAs(Path.Combine(Server.MapPath("~/Photos/"), fileName));
                db.Products.Add(product);
                db.SaveChanges();
            }
            return View();
        }*/
    }
}
