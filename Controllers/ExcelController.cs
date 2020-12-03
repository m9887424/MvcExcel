using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using System.Web.Configuration;
using System.Net.Mime;
using ClosedXML.Excel;
using LinqToExcel;
using PagedList;
using MvcExcel.Models;


namespace MvcExcel.Controllers
{
    public class ExcelController : Controller
    {
        public static string fullFilePath; //定義完整的路徑
        public static bool flag = false; //判斷是否有上傳成功
        public ZipCodeDBEntities db = new ZipCodeDBEntities();
        public ActionResult Index(int page = 1)
        {

            int currentPage = page < 1 ? 1 : page;

            var qry = db.ZipCode.OrderBy(x => x.Sequence).ThenBy(x => x.Id).ThenBy(x => x.Zip);

            var result = qry.ToPagedList(currentPage, 10);

            if (TempData["path"] != null)
            {
                ViewBag.path = fullFilePath;
            }
            return View(result);
        }


        [HttpPost]
        public ActionResult Uploads(HttpPostedFileBase file)
        {
            if (file != null)
            {
                if (file.ContentLength > 0)
                {
                    var extension = Path.GetExtension(file.FileName);
                    var fileSavePath = WebConfigurationManager.AppSettings["UploadPath"];
                    if (extension == ".xls" || extension == ".xlsx")
                    {
                        // 更改檔名為當天日期
                        var newFileName = string.Concat(DateTime.Now.ToString("yyyy-MM-dd HH-mm"), extension.ToLower());
                        fullFilePath = Path.Combine(Server.MapPath(fileSavePath), newFileName);
                        // 存放檔案到伺服器上
                        file.SaveAs(fullFilePath);
                        // 將資料路徑傳送到前端顯示
                        TempData["path"] = fullFilePath;

                        flag = true;

                        TempData["message"] = "檔案上傳成功";

                        return RedirectToAction("Index");

                    }
                    TempData["message"] = "請上傳 .xls  或 .xlsx 格式的檔案";
                    return RedirectToAction("Index");
                }

            }
            TempData["message"] = "請選擇檔案";

            return RedirectToAction("Index");

        }

        [HttpPost]
        public ActionResult Downloads()
        {
            if (flag)
            {
                FileInfo fl = new FileInfo(fullFilePath);
                //指定Content-Disposition(告訴用戶端瀏覽器如何處理附加文件)
                var cd = new ContentDisposition
                {
                    FileName = fl.Name,
                    Inline = false,
                    // Inline 設為 false 表示不要在瀏覽器上打開 
                };
                Response.AppendHeader("Content-Disposition", cd.ToString());
                Response.BufferOutput = false;
                //若下載的檔案太大, 需要將Response.BufferOutput 設為false
                //不然由於IIS的限制,可能會讓我們遇到 Overflow or underflow in the arithmetic operation的錯誤訊息
                var readStream = new FileStream(fl.FullName, FileMode.Open, FileAccess.Read);
                //指定Content-Type (告訴用戶端瀏覽器欲回應的內容)
                string contentType = MimeMapping.GetMimeMapping(fl.FullName);
                return File(readStream, contentType);
            }
            else
            {
                TempData["message"] = "請先上傳檔案";
                return RedirectToAction("Index");
            }
        }

        /*
        // 匯入Excel by ClosedXML
        [HttpPost]
        public ActionResult Imports()
        {
            //判斷上傳的檔案是否存在
            if (fullFilePath != null)
            {
                // “\”是一個轉義字符，所以需要用兩個代表一個
                // 路徑類似 D:\\temp\\Tim.xlsx
                var file_paths = fullFilePath.ToString().Replace("\\", "\\\\");
                XLWorkbook workbook = new XLWorkbook(file_paths);

                //讀取第一個Sheet
                IXLWorksheet worksheet = workbook.Worksheet(1);

                // 定義資料起始/結束 Cell
                var firstCell = worksheet.FirstCellUsed();
                var lastCell = worksheet.LastCellUsed();

                // 使用資料起始/結束 Cell，來定義出一個資料範圍
                var data = worksheet.Range(firstCell.Address, lastCell.Address);

                // 將資料範圍轉型
                var table = data.AsTable();

                //讀取資料，讀取的對象為 row:1 / column:1
                //string Excel = "";
                //Excel = table.Cell(1, 1).Value.ToString();

                //寫入資料
                table.Cell(2, 1).Value = "test";


                //資料顯示
                //Response.Write("<script language=javascript>alert('" + Excel + "');</" + "script>");

                if (System.IO.File.Exists(fullFilePath))
                {
                    workbook.SaveAs(file_paths);
                }
                else
                {
                    Response.Write("<script language=javascript>alert('請先上傳檔案');</" + "script>");
                }
            }
            return View("Index");
        }
        */
        [HttpPost]
        public ActionResult Imports()
        {
            var importZipCodes = new List<ZipCode>();
            if (flag)
            {
                var excelFile = new ExcelQueryFactory(fullFilePath);
                //欄位對映
                excelFile.AddMapping<ZipCode>(x => x.Id, "ID");
                excelFile.AddMapping<ZipCode>(x => x.Zip, "Zip");
                excelFile.AddMapping<ZipCode>(x => x.City, "CityName");
                excelFile.AddMapping<ZipCode>(x => x.Town, "Town");
                excelFile.AddMapping<ZipCode>(x => x.Sequence, "Sequence");

                //SheetName
                var excelContent = excelFile.Worksheet<ZipCode>("臺灣郵遞區號");

                //檢查資料
               
                foreach (var row in excelContent)
                {
                    var zipCode = new ZipCode();
                    zipCode.Id = row.Id;
                    zipCode.Sequence = row.Sequence;
                    zipCode.Zip = row.Zip;
                    zipCode.City = row.City;
                    zipCode.Town = row.Town;

                    importZipCodes.Add(zipCode);
                }

                //先砍掉全部資料
                foreach (var item in db.ZipCode.OrderBy(x => x.Id))
                {
                    db.ZipCode.Remove(item);
                }
                db.SaveChanges();

                //再把匯入的資料給存到資料庫
                foreach (var item in importZipCodes)
                {
                    db.ZipCode.Add(item);
                }
                db.SaveChanges();

                TempData["message"] = "匯入完成";
                return RedirectToAction("Index");
            }
            else
            {
                TempData["message"] = "請先上傳檔案";
                return RedirectToAction("Index");
            }

        }
    }
}
