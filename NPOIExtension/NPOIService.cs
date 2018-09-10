using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web; 

namespace NPOIExtension
{
    /// <summary>
    /// 基础类定义
    /// </summary>
    public class ExcelColumn
    {
        public string Name { get; set; }

        public string Field { get; set; }

        public int Width { get; set; }

        public int Colspan { get; set; }

        public int Rowspan { get; set; }

        public bool AllowNull { get; set; }

        public ICellStyle HeadCellStyle { get; set; }

        public ICellStyle DataCellStyle { get; set; }

        public Action<ICell, ISheet, IWorkbook> OnDataCellCreate { get; set; }

        public Action<ICell, ISheet, IWorkbook> OnHeadCellCreate { get; set; }

        public Func<object, object, ICell, object> OnDataBind { get; set; }

        public string CellFormulaString { get; set; }

        /// <summary>
        /// 用于个性化配置，一但赋值，不得随意更改
        /// </summary>
        public int ProfileCode { get; set; }

        public string Remark { get; set; }
    }

    /// <summary>
    /// 泛型
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExcelColumn<T> : ExcelColumn
    {
        public new Func<object, T, ICell, object> OnDataBind { get; set; }

        public Expression<Func<T, object>> FieldExpr { get; set; }
    }

    public class ExcelTitle
    {
        public string Text { get; set; }

        public ICellStyle Style { get; set; }

        private int _rowCount = 1;
        public int RowCount
        {
            get
            {
                return _rowCount;
            }
            set
            {
                _rowCount = value;
            }
        }

        public int ColumnCount { get; set; }
    }

    public static class NPOIExtension
    {
        public static int GetCellNum(this IRow row, string name)
        {
            for (int i = 0; i < row.LastCellNum; i++)
            {
                if (NPOIService.GetCellValue(row.GetCell(i)).ToString().Trim().Trim('`') == name.Trim().Trim('`'))
                {
                    return i;
                }
            }
            return -1;
        }

        public static int AddMergedRegion(this ISheet sheet, CellRangeAddress region, bool isMergeData)
        {
            int index = sheet.AddMergedRegion(region);
            if (isMergeData == true)
            {
                for (int i = region.FirstRow + 1; i <= region.LastRow; i++)
                {
                    for (int j = region.FirstColumn; j <= region.LastColumn; j++)
                    {
                        sheet.GetRow(i).GetCell(j).SetCellValue("");
                    }
                }
            }
            return index;
        }
    }

    /// <summary>
    /// NPOI服务类
    /// </summary>
    public class NPOIService
    {
        /// <summary>
        /// 针对web端创建IWorkbook
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        public static IWorkbook GetWorkbook(HttpRequestBase request)
        {
            if (request.Files.Count == 0)
            {
                throw new ExcelValidException("请选择文件");
            }
            var fileUpload = request.Files[0];
            string fileExtension = Path.GetExtension(fileUpload.FileName).ToLower();
            if (fileExtension != ".xls" && fileExtension != ".xlsx")
            {
                throw new ExcelValidException("选择的文件必须是Excel格式的，请重新选择！");
            }
            IWorkbook workbook = null;
            if (fileExtension == ".xls")
            {
                workbook = new HSSFWorkbook(fileUpload.InputStream);
            }
            else
            {
                workbook = new XSSFWorkbook(fileUpload.InputStream);
            }
            return workbook;
        }

        public static object GetCellValue(ICell cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }
            switch (cell.CellType)
            {
                case CellType.Blank:
                    return string.Empty;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue;
                    }
                    return cell.NumericCellValue;
                case CellType.String:
                    return (cell.StringCellValue ?? string.Empty).Trim().Trim('`');
                case CellType.Error:
                    return cell.ErrorCellValue;
                case CellType.Formula:
                    var evaluator = cell.Row.Sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator();
                    return GetCellValue(evaluator.EvaluateInCell(cell));
                default:
                    return cell.StringCellValue ?? "";
            }
        }

        /// <summary>
        /// 把列表导入Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="workbook"></param>
        /// <param name="sheet"></param>
        /// <param name="collection"></param>
        /// <param name="columns"></param>
        /// <param name="title"></param>
        public static void AppendToExcel<T>(IWorkbook workbook, ISheet sheet, List<T> collection, List<ExcelColumn> columns, int startRow, ExcelTitle title = null)
        {
            AppendToExcelProcess<T>(workbook, sheet, collection, columns, startRow, title);
        }

        /// <summary>
        /// 把列表导入Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="workbook"></param>
        /// <param name="sheet"></param>
        /// <param name="collection"></param>
        /// <param name="columns"></param>
        /// <param name="title"></param>
        public static void ListToExcel<T>(IWorkbook workbook, ISheet sheet, List<T> collection, List<ExcelColumn> columns, ExcelTitle title = null)
        {
            ListToExcelProcess<T>(workbook, sheet, collection, columns, title);
        }

        /// <summary>
        /// 把列表导入Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="workbook"></param>
        /// <param name="sheet"></param>
        /// <param name="collection"></param>
        /// <param name="columns"></param>
        /// <param name="title"></param>
        public static void AppendToExcelProcess<T>(IWorkbook workbook, ISheet sheet, List<T> collection, List<ExcelColumn> columns, int startRow, ExcelTitle title = null)
        {
            int titleRowCount = 0;
            // 标题
            if (title != null)
            {
                titleRowCount = title.RowCount;
                title.ColumnCount = (title.ColumnCount == 0 ? columns.Count : title.ColumnCount);
                CreateTitle(workbook, sheet, title, 0, title.ColumnCount - 1);
            }
            // 初始化表头
            InitHead<T>(columns);
            // 数据绑定
            ListToExcelDataRow(workbook, sheet, collection, columns, startRow);
        }

        /// <summary>
        /// 把列表导入Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="workbook"></param>
        /// <param name="sheet"></param>
        /// <param name="collection"></param>
        /// <param name="columns"></param>
        /// <param name="title"></param>
        public static void ListToExcelProcess<T>(IWorkbook workbook, ISheet sheet, List<T> collection, List<ExcelColumn> columns, ExcelTitle title = null)
        {
            int titleRowCount = 0;
            // 标题
            if (title != null)
            {
                titleRowCount = title.RowCount;
                title.ColumnCount = (title.ColumnCount == 0 ? columns.Count : title.ColumnCount);
                CreateTitle(workbook, sheet, title, 0, title.ColumnCount - 1);
            }

            // 初始化表头
            InitHead<T>(columns);

            // 表头
            CreateHead(workbook, sheet, columns, titleRowCount);

            // 数据绑定
            ListToExcelDataRow(workbook, sheet, collection, columns, titleRowCount + 1);
        }

        public static List<T> ExcelToList<T>(IWorkbook workbook, ISheet sheet, List<ExcelColumn> columns, int headRowIndex, int dataRowIndex)
        {
            InitHead<T>(columns);
            List<int> lstCellNum = new List<int>();
            List<T> entityList = new List<T>();
            IRow headRow = sheet.GetRow(headRowIndex);
            for (int i = 0; i < columns.Count; i++)
            {
                int cellNum = headRow.GetCellNum(columns[i].Name);
                if (cellNum == -1)
                {
                    throw new ExcelValidException("缺少列：" + columns[i].Name);
                }
                lstCellNum.Add(cellNum);
            }
            PropertyInfo[] properties = typeof(T).GetProperties();
            Dictionary<string, PropertyInfo> dictProperty = new Dictionary<string, PropertyInfo>();
            for (int i = 0; i < properties.Length; i++)
            {
                dictProperty.Add(properties[i].Name, properties[i]);
            }
            IRow row;
            ICell cell;
            object cellValue;
            for (int i = dataRowIndex; i <= sheet.LastRowNum; i++)
            {
                row = sheet.GetRow(i);
                T entity = Activator.CreateInstance<T>();
                for (int j = 0; j < columns.Count; j++)
                {
                    cell = row.GetCell(lstCellNum[j]);
                    cellValue = GetCellValue(cell);
                    if (columns[j].AllowNull == false && string.IsNullOrWhiteSpace(cellValue.ToString()))
                    {
                        throw new ExcelValidException(columns[j].Name + "不能为空");
                    }
                    if (columns[j] is ExcelColumn<T>)
                    {
                        var genericColumn = columns[j] as ExcelColumn<T>;
                        if (genericColumn.OnDataBind != null)
                        {
                            cellValue = genericColumn.OnDataBind(cellValue, entity, cell);
                        }
                    }
                    else
                    {
                        if (columns[j].OnDataBind != null)
                        {
                            cellValue = columns[j].OnDataBind(cellValue, entity, cell);
                        }
                    }
                    if (dictProperty[columns[j].Field].PropertyType.IsValueType && string.IsNullOrWhiteSpace(cellValue.ToString()))
                    {
                        cellValue = Activator.CreateInstance(dictProperty[columns[j].Field].PropertyType);
                    }
                    else
                    {
                        try
                        {
                            cellValue = Convert.ChangeType(cellValue, dictProperty[columns[j].Field].PropertyType);
                        }
                        catch
                        {
                            if (dictProperty[columns[j].Field].PropertyType == typeof(DateTime))
                            {
                                cellValue = CommonUtility.ConvertToDateTime(cellValue.ToString());
                            }
                        }
                    }
                    dictProperty[columns[j].Field].SetValue(entity, cellValue, null);

                }
                entityList.Add(entity);
            }
            return entityList;
        }

        public static List<T> ExcelToList<T>(IWorkbook workbook, ISheet sheet, List<ExcelColumn> columns, int headRowIndex = 0)
        {
            return ExcelToList<T>(workbook, sheet, columns, headRowIndex, headRowIndex + 1);
        }

        public static void InitHead<T>(List<ExcelColumn> columns)
        {
            PropertyInfo[] properties = typeof(T).GetProperties();
            for (int i = 0; i < columns.Count; i++)
            {
                if (columns[i] is ExcelColumn<T>)
                {
                    var genericColumn = columns[i] as ExcelColumn<T>;
                    if (genericColumn.FieldExpr != null)
                    {
                        var memberExpression = genericColumn.FieldExpr.Body as MemberExpression;
                        if (memberExpression == null)
                        {
                            memberExpression = (genericColumn.FieldExpr.Body as UnaryExpression).Operand as MemberExpression;
                        }
                        columns[i].Field = memberExpression.Member.Name;
                    }
                    if (!string.IsNullOrWhiteSpace(columns[i].Name))
                    {
                        continue;
                    }
                    var property = properties.FirstOrDefault(u => u.Name == columns[i].Field);
                    if (property == null)
                    {
                        continue;
                    }
                    DisplayNameAttribute[] attributes = (DisplayNameAttribute[])property.GetCustomAttributes(typeof(DisplayNameAttribute), true);
                    if (attributes != null && attributes.Length > 0)
                    {
                        columns[i].Name = attributes[0].DisplayName;
                    }
                }
            }
        }

        private static void CreateHead(IWorkbook workbook, ISheet sheet, List<ExcelColumn> columns, int startRow)
        {
            ICell cell;
            IRow row;
            row = sheet.CreateRow(startRow);
            for (int i = 0; i < columns.Count; i++)
            {
                if (columns[i].Width > 0)
                {
                    sheet.SetColumnWidth(i, 20 * columns[i].Width);
                }
                cell = row.CreateCell(i);
                cell.CellStyle = columns[i].HeadCellStyle == null ? GetCellStyle(workbook) : columns[i].HeadCellStyle;
                if (columns[i].OnHeadCellCreate != null)
                {
                    columns[i].OnHeadCellCreate(cell, sheet, workbook);
                }
                cell.SetCellValue(columns[i].Name);
            }
            // sheet.CreateFreezePane(0, startRow + 1, 0, startRow + 1); 冻结表头

            for (int i = 0; i < columns.Count; i++)
            {
                if (columns[i].Colspan > 1)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(startRow, startRow, i, i + columns[i].Colspan - 1));
                }
            }
        }

        private static void CreateTitle(IWorkbook workbook, ISheet sheet, ExcelTitle title, int startColumn, int endColumn)
        {
            ICell cell;
            IRow row;
            int titleRowCount = 0;
            if (title != null)
            {
                titleRowCount = title.RowCount;
                if (title.Text != null)
                {
                    row = sheet.CreateRow(0);
                    cell = row.CreateCell(0);
                    if (title.Style == null)
                    {
                        title.Style = GetTitleStyle(workbook);
                    }
                    cell.CellStyle = title.Style;
                    cell.SetCellValue(title.Text);
                    sheet.AddMergedRegion(new CellRangeAddress(0, title.RowCount - 1, startColumn, endColumn));
                }
            }
        }

        /// <summary>
        ///填充Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="workbook"></param>
        /// <param name="sheet"></param>
        /// <param name="collection"></param>
        /// <param name="columns"></param>
        /// <param name="startRow"></param>
        private static void ListToExcelDataRow<T>(IWorkbook workbook, ISheet sheet, List<T> collection, List<ExcelColumn> columns, int startRow)
        {
            ICellStyle defaultCellStyle = GetCellStyle(workbook);
            ICell cell;
            IRow row;

            List<PropertyInfo> lstProperty = new List<PropertyInfo>();
            Type type = typeof(T);
            for (int i = 0; i < columns.Count; i++)
            {
                if (string.IsNullOrEmpty(columns[i].Field))
                {
                    lstProperty.Add(null);
                }
                else
                {
                    lstProperty.Add(type.GetProperty(columns[i].Field));
                }
            }

            for (int i = 0; i < collection.Count; i++)
            {
                row = sheet.CreateRow(i + startRow);
                for (int j = 0; j < lstProperty.Count; j++)
                {
                    cell = row.CreateCell(j);
                    cell.CellStyle = columns[j].DataCellStyle == null ? defaultCellStyle : columns[j].DataCellStyle;

                    if (!string.IsNullOrWhiteSpace(columns[j].CellFormulaString))
                    {
                        cell.SetCellFormula(columns[j].CellFormulaString);
                    }
                    if (columns[j].OnDataCellCreate != null)
                    {
                        columns[j].OnDataCellCreate(cell, sheet, workbook);
                    }
                    object value = null;
                    if (lstProperty[j] != null)
                    {
                        value = lstProperty[j].GetValue(collection[i], null);
                    }
                    if (columns[j] is ExcelColumn<T>)
                    {
                        var genericColumn = columns[j] as ExcelColumn<T>;
                        if (genericColumn.OnDataBind != null)
                        {
                            value = genericColumn.OnDataBind(value, collection[i], cell);
                        }
                    }
                    else
                    {
                        if (columns[j].OnDataBind != null)
                        {
                            value = columns[j].OnDataBind(value, collection[i], cell);
                        }
                    }
                    if (value == null)
                    {
                        continue;
                    }
                    TypeCode typeCode = Type.GetTypeCode(value.GetType());
                    if (typeCode == TypeCode.Decimal && columns[j].DataCellStyle == null)
                    {
                        if (cell.CellStyle.Index == defaultCellStyle.Index)
                        {
                            cell.CellStyle = defaultCellStyle;
                        }
                    }
                    if (typeCode == TypeCode.DateTime)
                    {
                        #region 过滤默认时间 1900-01-01
                        try
                        {
                            if (Convert.ToDateTime(value).Date <= Convert.ToDateTime("1900-01-01").Date)
                                cell.SetCellValue("");
                            else cell.SetCellValue(Convert.ToDateTime(value).ToString("yyyy-MM-dd HH:mm:ss"));
                        }
                        catch
                        {
                            cell.SetCellValue("格式错误");
                        }
                        #endregion
                    }
                    else if (typeCode == TypeCode.Boolean)
                    {
                        #region bool类型：是/否
                        try
                        {
                            if (Convert.ToBoolean(value))
                                cell.SetCellValue("是");
                            else cell.SetCellValue("否");
                        }
                        catch
                        {
                            cell.SetCellValue("未知");
                        }
                        #endregion
                    }
                    else if (typeCode == TypeCode.Byte || typeCode == TypeCode.Decimal || typeCode == TypeCode.Double || typeCode == TypeCode.Int16 || typeCode == TypeCode.Int32 || typeCode == TypeCode.Int64 || typeCode == TypeCode.SByte || typeCode == TypeCode.Single || typeCode == TypeCode.UInt16 || typeCode == TypeCode.UInt32 || typeCode == TypeCode.UInt64)
                    {
                        cell.SetCellValue(Convert.ToDouble(value));
                    }
                    else
                    {
                        cell.SetCellValue(value.ToString());
                    }
                }
            }
        }

        #region 样式设置，可自定义自己的样式

        /// <summary>
        /// 设置默认单元格样式
        /// </summary>
        /// <param name="hssfWorkbook">HSSFWorkbook</param>
        /// <returns>cellStyle</returns>
        public static ICellStyle GetCellStyle(IWorkbook workBook)
        {
            ICellStyle cellStyle = workBook.CreateCellStyle();
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.SetFont(GetFont(workBook));
            return cellStyle;
        }

        public static IFont GetFont(IWorkbook workBook)
        {
            IFont font = workBook.CreateFont();
            font.FontHeightInPoints = 10;
            font.FontName = "Arial";
            return font;
        }

        public static ICellStyle GetTitleStyle(IWorkbook workBook)
        {
            ICellStyle style = workBook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            IFont font = workBook.CreateFont();
            font.FontHeight = 17 * 17;
            style.SetFont(font);
            return style;
        }

        /// <summary>
        /// 标题行样式
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="FontHeight"></param>
        /// <returns></returns>
        public static ICellStyle GetGreenTitleStyle(IWorkbook workBook, double FontHeight, bool addColor)
        {
            ICellStyle style = workBook.CreateCellStyle();//设置标题行样式
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            if (addColor)
            {
                style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;//背景色
                style.FillPattern = FillPattern.Squares;
                style.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;//前景色
                IFont font = workBook.CreateFont();
                font.FontHeight = FontHeight;//字体大小
                font.Boldweight = (short)FontBoldWeight.Bold;//字体加粗
                style.SetFont(font);
            }
            return style;
        }
        /// <summary>
        /// 行标题样式
        /// </summary>
        /// <param name="workBook"></param>
        /// <returns></returns>
        public static ICellStyle GetLightCornflowerBlueTitleStyle(IWorkbook workBook)
        {
            ICellStyle titleStyle = workBook.CreateCellStyle();//设置标题行样式
            titleStyle.Alignment = HorizontalAlignment.Center;
            titleStyle.VerticalAlignment = VerticalAlignment.Center;
            titleStyle.BorderBottom = BorderStyle.Thin;
            titleStyle.BorderTop = BorderStyle.Thin;
            titleStyle.BorderLeft = BorderStyle.Thin;
            titleStyle.BorderRight = BorderStyle.Thin;
            titleStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightCornflowerBlue.Index;
            titleStyle.FillPattern = FillPattern.Squares;
            titleStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.LightCornflowerBlue.Index;
            IFont titleFont = workBook.CreateFont();
            titleFont.FontHeight = 15 * 15;
            titleFont.Boldweight = (short)FontBoldWeight.Bold;
            titleStyle.SetFont(titleFont);
            return titleStyle;
        }

        /// <summary>
        /// 内容行样式(黄色背景红色字体)
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="FontHeight"></param>
        /// <returns></returns>
        public static ICellStyle GetYellowTitleStyle(IWorkbook workBook)
        {
            ICellStyle style = workBook.CreateCellStyle();//设置标题行样式
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Top;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;//背景色
            style.FillPattern = FillPattern.Squares;
            style.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;//前景色
            IFont font = workBook.CreateFont();
            font.FontHeight = 14 * 14;//字体大小
            font.Color = NPOI.HSSF.Util.HSSFColor.Red.Index;
            //font.Boldweight = (short)FontBoldWeight.Bold;//字体加粗
            style.SetFont(font);

            return style;
        }

        #endregion
    }

    /// <summary>
    /// Excel 参数异常校验
    /// </summary>
    public class ExcelValidException : Exception
    {
        public ExcelValidException(string message) : base(message) { }

        public ExcelValidException(string message, Exception innerException) : base(message, innerException) { }
    }
}
