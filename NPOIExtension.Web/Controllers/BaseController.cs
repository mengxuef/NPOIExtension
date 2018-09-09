using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOIExtension.Web.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NPOIExtension.Web.Controllers
{
    public class BaseController : Controller
    {
        /// <summary>
        /// 下载Excel2007到页面
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="hssfWorkbook"></param>
        /// <returns></returns>
        public ActionResult Excel2007File(string fileName, IWorkbook workbook)
        {
            return new NPOIFileResult(workbook, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") { FileDownloadName = fileName + ".xlsx" };
        }
        /// <summary>
        /// 下载Excel2003到页面
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="hssfWorkbook"></param>
        /// <returns></returns>
        public ActionResult Excel2003File(string fileName, IWorkbook workbook)
        {
            return new NPOIFileResult(workbook, "application/vnd.ms-excel") { FileDownloadName = fileName + ".xls" };
        }

        public ActionResult ExportExcel(string fileName, IWorkbook workbook)
        {
            if (workbook.GetType() == typeof(HSSFWorkbook))
            {
                return Excel2003File(fileName, workbook);
            }
            else
            {
                return Excel2007File(fileName, workbook);
            }
        }
        /// <summary>
        /// 根据ExcelVersion 创建IWorkbook
        /// </summary>
        /// <returns></returns>
        public IWorkbook CreateWorkbook()
        {
            string excelVersion = Request["ExcelVersion"];
            if (excelVersion == "2003")
            {
                return new HSSFWorkbook();
            }
            else if (excelVersion == "2007")
            {
                return new XSSFWorkbook();
            }
            return new HSSFWorkbook();
        }


    }
}