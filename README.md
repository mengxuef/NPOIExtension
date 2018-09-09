# NPOI--基于NPOI封装高级应用

只需要引用NPOI官网相关dll，然后把NPOIService 和CommonUtility 引用到项目中即可,可以快速构建项目导出，无需关注NPOI底层实现。

# 文件说明
NPOIService：基于NPOI分装的EXCEL主要核心代码
  
  # 核心方法
 
  1、把List封装为Excel
  ~~~C#
  public static void ListToExcel<T>(IWorkbook workbook, ISheet sheet, List<T> collection, List<ExcelColumn> columns, ExcelTitle title = null)
  ~~~
  2、Excel分装为List
  ~~~C#
  public static List<T> ExcelToList<T>(IWorkbook workbook, ISheet sheet, List<ExcelColumn> columns, int headRowIndex = 0)
   ~~~
  3、读取Excel模板 填充数据
  ~~~C#
  public static void AppendToExcel<T>(IWorkbook workbook, ISheet sheet, List<T> collection, List<ExcelColumn> columns, int startRow, ExcelTitle title = null)
  ~~~
  4、数据绑定
  ~~~C#
  private static void ListToExcelDataRow<T>(IWorkbook workbook, ISheet sheet, List<T> collection, List<ExcelColumn> columns, int startRow)
  ~~~
  
  # MVC中案例
    导出Excel
  ~~~C#
  public ActionResult SettlementExport()
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
  ~~~
  
  导入Excel
  
  ~~~C#
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
  ~~~
  
  
