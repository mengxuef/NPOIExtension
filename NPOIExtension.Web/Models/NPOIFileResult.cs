using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NPOIExtension.Web.Models
{
    public class NPOIFileResult : FileResult
    {
        public IWorkbook Workbook { get; private set; }

        public NPOIFileResult(IWorkbook workbook,string contentType) : base(contentType)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException("workbook");
            }
            this.Workbook = workbook;
        }

        protected override void WriteFile(HttpResponseBase response)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            Workbook.Write(response.OutputStream);
            sw.Stop();
            response.AddHeader("Export-Output", sw.ElapsedMilliseconds.ToString());
        }
    }
}