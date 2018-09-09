using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOIExtension.Web.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NPOIExtension.Web.Controllers
{
    public class HomeController : BaseController
    {
        List<SettlementModel> settlementList = null;
        public HomeController()
        {
            settlementList = new List<SettlementModel>();
            settlementList.Add(new SettlementModel() { Id = 1, BussinessType = 1, CreateDateBegin = DateTime.Now, CreateDateEnd = DateTime.Now.AddDays(-30), SerialId = "SZ1001", Title = "XXX月结对账单" });
            settlementList.Add(new SettlementModel() { Id = 2, BussinessType = 2, CreateDateBegin = DateTime.Now, CreateDateEnd = DateTime.Now.AddDays(-30), SerialId = "SZ1001", Title = "XXX月结对账单" });
            settlementList.Add(new SettlementModel() { Id = 3, BussinessType = 1, CreateDateBegin = DateTime.Now, CreateDateEnd = DateTime.Now.AddDays(-30), SerialId = "SZ1002", Title = "XXX月结对账单" });
            settlementList.Add(new SettlementModel() { Id = 4, BussinessType = 2, CreateDateBegin = DateTime.Now, CreateDateEnd = DateTime.Now.AddDays(-30), SerialId = "SZ1003", Title = "XXX月结对账单" });
            settlementList.Add(new SettlementModel() { Id = 5, BussinessType = 1, CreateDateBegin = DateTime.Now, CreateDateEnd = DateTime.Now.AddDays(-30), SerialId = "SZ1004", Title = "XXX月结对账单" });
            settlementList.Add(new SettlementModel() { Id = 6, BussinessType = 2, CreateDateBegin = DateTime.Now, CreateDateEnd = DateTime.Now.AddDays(-30), SerialId = "SZ1005", Title = "XXX月结对账单" });
        }
        /// <summary>
        /// Excel导出demo
        /// </summary>
        /// <returns></returns>
        public ActionResult ExcelExport()
        {
            //1、获取数据
            var list = settlementList;
            //2、IWorkbook 和ISheet
            IWorkbook workbook = new HSSFWorkbook();//new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("对账单");

            //3、设置导出的列
            List<ExcelColumn> columns = new List<ExcelColumn>();
            columns.Add(new ExcelColumn<SettlementModel> { FieldExpr = u => u.Id });
            columns.Add(new ExcelColumn<SettlementModel>
            {
                FieldExpr = u => u.BussinessType,
                Name = "业务类型2",
                OnDataBind = (value, item, cell) => item.BussinessType == 1 ? "月结对账单" : "季度结对账单"
            });
            columns.Add(new ExcelColumn<SettlementModel> { FieldExpr = u => u.CreateDateBegin, Width = 100 });
            columns.Add(new ExcelColumn<SettlementModel> { FieldExpr = u => u.CreateDateEnd });
            columns.Add(new ExcelColumn<SettlementModel> { FieldExpr = u => u.SerialId });
            columns.Add(new ExcelColumn<SettlementModel> { FieldExpr = u => u.Title });
            //4、把数据和列做绑定
            NPOIService.ListToExcel(workbook, sheet, list, columns);
            //5、导出excel 
            return ExportExcel(DateTime.Now.ToString("yyyyMMddHHmmss"), workbook);

        }

        /// <summary>
        /// Excel导入案例
        /// </summary>
        /// <returns></returns>
        public ActionResult ExcelImport()
        {
            try
            {
                IWorkbook workbook = NPOIService.GetWorkbook(HttpContext.Request);
                ISheet sheet = workbook.GetSheetAt(0);

                List<ExcelColumn> columns = new List<ExcelColumn>();
                columns.Add(new ExcelColumn<SettlementModel> { FieldExpr = u => u.Id });
                columns.Add(new ExcelColumn<SettlementModel> { FieldExpr = u => u.BussinessType });
                columns.Add(new ExcelColumn<SettlementModel> { FieldExpr = u => u.CreateDateBegin, AllowNull = true });
                columns.Add(new ExcelColumn<SettlementModel> { FieldExpr = u => u.CreateDateEnd });
                columns.Add(new ExcelColumn<SettlementModel> { FieldExpr = u => u.SerialId });
                columns.Add(new ExcelColumn<SettlementModel> { FieldExpr = u => u.Title });

                var list = NPOIService.ExcelToList<SettlementModel>(workbook, sheet, columns);
                //业务逻辑处理

            }
            catch (Exception ex)
            {
                //异常处理
            }
            return Json(new { Success = true }, JsonRequestBehavior.AllowGet);
        }
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}