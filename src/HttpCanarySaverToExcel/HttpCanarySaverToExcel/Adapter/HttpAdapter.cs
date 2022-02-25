using LitJson;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace HttpCanarySaverToExcel.Adapter
{
    public class HttpAdapter : AdapterBase
    {
        /**
         * 请求部分
         */
        // 除去ueryString的url
        public string ReqtUrl { get; set; }
        public string ReqMethod { get; set; }
        public string ReqContentType { get; set; }
        // url中的queryString
        public Dictionary<string, string> ReqQueryStringDict = new Dictionary<string, string>();
        public List<Cookie> ReqCookies { get; set; } = new List<Cookie>();
        // key:一些固定分析的字段名全小写, value:实际抓包到的字段名（区分大小写）
        public Dictionary<string, string> ReqHeaderKeyToRealName = new Dictionary<string, string>();
        // 除去cookie等固定分析的字段以及要忽略的字段外的其他header字段
        public Dictionary<string, string> ReqHeaders = new Dictionary<string, string>();
        // 如果请求为json形式
        public string ReqJsonParam { get; set; }
        // 如果请求为表单形式
        public Dictionary<string, string> ReqFormParamDict { get; set; }
        /**
         * 服务器响应部分
         */
        public string RespContentType { get; set; }
        public int RespStateCode { get; set; }
        public List<Cookie> RespCookies { get; set; } = new List<Cookie>();
        // key:一些固定分析的字段名全小写, value:实际抓包到的字段名（区分大小写）
        public Dictionary<string, string> RespHeaderKeyToRealName = new Dictionary<string, string>();
        // 除去cookie等固定分析的字段以及要忽略的字段外的其他header字段
        public Dictionary<string, string> RespHeaders = new Dictionary<string, string>();
        // 如果响应为json形式
        public string RespJson { get; set; }
        // 如果响应为文本类型
        public string RespText { get; set; }

        public override void WriteExcelSheet(IWorkbook workbook, UserConfig userConfig, OneOverviewRecord overviewRecord)
        {
            int seq = overviewRecord.RecordFileInfo.NumberSeq;
            ISheet sheet = workbook.CreateSheet(seq.ToString());
            // 设置2列的宽度
            for (int columnIndex = 0; columnIndex < AppConst.DETIAL_SHEET_COLUMN_WIDTH.Length; columnIndex++)
                sheet.SetColumnWidth(columnIndex, (int)(AppConst.DETIAL_SHEET_COLUMN_WIDTH[columnIndex] * 256f));

            /**
             * 本Excel中固定的格式
             */
            // 普通格式（所有单元格都设置内外边框，文字垂直方向居上，水平方向居左）
            ICellStyle NORMAL_CELL_STYLE = workbook.CreateCellStyle();
            NORMAL_CELL_STYLE.BorderTop = BorderStyle.Thin;
            NORMAL_CELL_STYLE.BorderBottom = BorderStyle.Thin;
            NORMAL_CELL_STYLE.BorderLeft = BorderStyle.Thin;
            NORMAL_CELL_STYLE.BorderRight = BorderStyle.Thin;
            NORMAL_CELL_STYLE.VerticalAlignment = VerticalAlignment.Top;
            NORMAL_CELL_STYLE.Alignment = HorizontalAlignment.Left;
            // 加粗字体的单元格格式
            ICellStyle BOLD_CELL_STYLE = workbook.CreateCellStyle();
            BOLD_CELL_STYLE.BorderTop = BorderStyle.Thin;
            BOLD_CELL_STYLE.BorderBottom = BorderStyle.Thin;
            BOLD_CELL_STYLE.BorderLeft = BorderStyle.Thin;
            BOLD_CELL_STYLE.BorderRight = BorderStyle.Thin;
            BOLD_CELL_STYLE.VerticalAlignment = VerticalAlignment.Top;
            BOLD_CELL_STYLE.Alignment = HorizontalAlignment.Left;
            IFont boldFont = workbook.CreateFont();
            boldFont.IsBold = true;
            BOLD_CELL_STYLE.SetFont(boldFont);
            // 支持单元格中有换行的单元格格式
            ICellStyle WRAP_TEXT_CELL_STYLE = workbook.CreateCellStyle();
            WRAP_TEXT_CELL_STYLE.WrapText = true;
            WRAP_TEXT_CELL_STYLE.BorderTop = BorderStyle.Thin;
            WRAP_TEXT_CELL_STYLE.BorderBottom = BorderStyle.Thin;
            WRAP_TEXT_CELL_STYLE.BorderLeft = BorderStyle.Thin;
            WRAP_TEXT_CELL_STYLE.BorderRight = BorderStyle.Thin;
            WRAP_TEXT_CELL_STYLE.VerticalAlignment = VerticalAlignment.Top;
            WRAP_TEXT_CELL_STYLE.Alignment = HorizontalAlignment.Left;
            // 超链接的单元格格式
            ICellStyle HYPERLINK_CELL_STYLE = workbook.CreateCellStyle();
            IFont HYPERLINK_FONT = workbook.CreateFont();
            HYPERLINK_FONT.Underline = FontUnderlineType.Single;
            HYPERLINK_FONT.Color = HSSFColor.Blue.Index;
            HYPERLINK_CELL_STYLE.SetFont(HYPERLINK_FONT);
            HYPERLINK_CELL_STYLE.BorderTop = BorderStyle.Thin;
            HYPERLINK_CELL_STYLE.BorderBottom = BorderStyle.Thin;
            HYPERLINK_CELL_STYLE.BorderLeft = BorderStyle.Thin;
            HYPERLINK_CELL_STYLE.BorderRight = BorderStyle.Thin;
            HYPERLINK_CELL_STYLE.VerticalAlignment = VerticalAlignment.Top;
            HYPERLINK_CELL_STYLE.Alignment = HorizontalAlignment.Left;

            int currentRowIndex = 0;

            // 第1行为返回总览超链接（点击跳转到总览Sheet表的这个条目）、序号
            {
                IRow row = sheet.CreateRow(currentRowIndex);
                ICell returnCell = row.CreateCell(0);
                returnCell.SetCellValue("返回总览");
                // 设置超链接到对应Sheet表
                XSSFHyperlink hyperlink = new XSSFHyperlink(HyperlinkType.Document);
                hyperlink.Address = ($"'{AppConst.OVERVIEW_SHEET_NAME}'!A{seq + AppConst.START_RECORD_ROW_INDEX}");
                returnCell.Hyperlink = hyperlink;
                returnCell.CellStyle = HYPERLINK_CELL_STYLE;

                ICell seqCell = row.CreateCell(1);
                seqCell.SetCellValue($"序号：{seq}");
                seqCell.CellStyle = NORMAL_CELL_STYLE;
            }
            currentRowIndex++;
            // 第2行为备注内容（通过公式引用总览页中填写的备注）
            {
                IRow row = sheet.CreateRow(currentRowIndex);
                row.CreateCell(0).SetCellValue("备注（引用自总览页）");
                row.CreateCell(1).SetCellFormula($"'{AppConst.OVERVIEW_SHEET_NAME}'!F{AppConst.START_RECORD_ROW_INDEX + seq}");
                row.GetCell(0).CellStyle = NORMAL_CELL_STYLE;
                row.GetCell(1).CellStyle = NORMAL_CELL_STYLE;
            }
            currentRowIndex++;
            // 第3行为存档文件路径的超链接
            {
                IRow row = sheet.CreateRow(currentRowIndex);
                row.CreateCell(0).SetCellValue("存档文件夹路径");
                row.GetCell(0).CellStyle = NORMAL_CELL_STYLE;
                ICell dirPathCell = row.CreateCell(1);
                dirPathCell.SetCellValue(overviewRecord.RecordFileInfo.FullFolderPath);
                // 设置超链接打开对应文件夹
                XSSFHyperlink hyperlink = new XSSFHyperlink(HyperlinkType.File);
                hyperlink.Address = overviewRecord.RecordFileInfo.FullFolderPath;
                dirPathCell.Hyperlink = hyperlink;
                dirPathCell.CellStyle = HYPERLINK_CELL_STYLE;
            }
            currentRowIndex++;
            // 第4行为“请求部分”的标题
            {
                // 标题行设置合并后居中
                IRow titleRow = sheet.CreateRow(currentRowIndex);
                titleRow.HeightInPoints = 22.5f;
                // 标题行水平居中
                ICellStyle titleStyle = workbook.CreateCellStyle();
                titleStyle.BorderTop = BorderStyle.Thin;
                titleStyle.BorderBottom = BorderStyle.Thin;
                titleStyle.BorderLeft = BorderStyle.Thin;
                titleStyle.BorderRight = BorderStyle.Thin;
                titleStyle.VerticalAlignment = VerticalAlignment.Top;
                titleStyle.Alignment = HorizontalAlignment.Center;
                // 标题行的字体
                IFont titleFont = workbook.CreateFont();
                titleFont.FontHeightInPoints = 18;
                titleFont.IsBold = true;
                titleStyle.SetFont(titleFont);
                // 由于NPOI的设计缺陷，要想对合并单元格设置边框，只能先将合并前的单元格进行边框设置后再合并
                for (int columnIndex = 0; columnIndex < 1; columnIndex++)
                {
                    // 这里必须先create出要被合并的单元格，否则NPOI获取不到也就无法设置单元格格式
                    titleRow.CreateCell(columnIndex).CellStyle = titleStyle;
                }

                ICell titleCell = titleRow.GetCell(0);
                titleCell.SetCellValue("请求部分");
                // 合并单元格，四个参数分别是合并起始行index、结束行index、合并起始列index、结束列index
                CellRangeAddress titleRegion = new CellRangeAddress(currentRowIndex, currentRowIndex, 0, 1);
                sheet.AddMergedRegion(titleRegion);
            }
            currentRowIndex++;
            // 第5行为URL
            {
                IRow row = sheet.CreateRow(currentRowIndex);
                row.CreateCell(0).SetCellValue("URL");
                row.CreateCell(1).SetCellValue(ReqtUrl);
                row.GetCell(0).CellStyle = NORMAL_CELL_STYLE;
                row.GetCell(1).CellStyle = NORMAL_CELL_STYLE;
            }
            currentRowIndex++;
            // 第6行为Method
            {
                IRow row = sheet.CreateRow(currentRowIndex);
                row.CreateCell(0).SetCellValue("Method");
                row.CreateCell(1).SetCellValue(ReqMethod);
                row.GetCell(0).CellStyle = NORMAL_CELL_STYLE;
                row.GetCell(1).CellStyle = NORMAL_CELL_STYLE;
            }
            currentRowIndex++;
            // 第7行为content-type
            {
                IRow row = sheet.CreateRow(currentRowIndex);
                row.CreateCell(0).SetCellValue(ReqHeaderKeyToRealName.ContainsKey("content-type") ? ReqHeaderKeyToRealName["content-type"] : "content-type");
                row.CreateCell(1).SetCellValue(string.IsNullOrEmpty(ReqContentType) ? "未声明" : ReqContentType);
                row.GetCell(0).CellStyle = NORMAL_CELL_STYLE;
                row.GetCell(1).CellStyle = NORMAL_CELL_STYLE;
            }
            currentRowIndex++;
            // 如果有QueryString，输出所有的keyValuePair
            if (ReqQueryStringDict.Count > 0)
            {
                IRow titleRow = sheet.CreateRow(currentRowIndex);
                titleRow.CreateCell(0).SetCellValue("QueryString如下，每行一个键值对");
                titleRow.GetCell(0).CellStyle = BOLD_CELL_STYLE;

                currentRowIndex++;
                foreach (var keyValuePair in ReqQueryStringDict)
                {
                    IRow row = sheet.CreateRow(currentRowIndex);
                    row.CreateCell(0).SetCellValue(keyValuePair.Key);
                    row.CreateCell(1).SetCellValue(keyValuePair.Value);
                    row.GetCell(0).CellStyle = NORMAL_CELL_STYLE;
                    row.GetCell(1).CellStyle = NORMAL_CELL_STYLE;

                    currentRowIndex++;
                }
            }
            // 如果有Cookie，输出所有的keyValuePair
            if (ReqCookies.Count > 0)
            {
                IRow titleRow = sheet.CreateRow(currentRowIndex);
                string cookieParamName = ReqHeaderKeyToRealName.ContainsKey("cookie") ? ReqHeaderKeyToRealName["cookie"] : "cookie";
                titleRow.CreateCell(0).SetCellValue($"{cookieParamName}如下，每行一个键值对");
                titleRow.GetCell(0).CellStyle = BOLD_CELL_STYLE;

                currentRowIndex++;
                foreach (Cookie cookie in ReqCookies)
                {
                    IRow row = sheet.CreateRow(currentRowIndex);
                    row.CreateCell(0).SetCellValue(cookie.Name);
                    row.CreateCell(1).SetCellValue(cookie.Value);
                    row.GetCell(0).CellStyle = NORMAL_CELL_STYLE;
                    row.GetCell(1).CellStyle = NORMAL_CELL_STYLE;

                    currentRowIndex++;
                }
            }
            // 如果有其他Header字段，输出所有的keyValuePair
            if (ReqHeaders.Count > 0)
            {
                IRow titleRow = sheet.CreateRow(currentRowIndex);
                titleRow.CreateCell(0).SetCellValue("其他Header字段，每行一个键值对");
                titleRow.GetCell(0).CellStyle = BOLD_CELL_STYLE;

                currentRowIndex++;
                foreach (var keyValuePair in ReqHeaders)
                {
                    IRow row = sheet.CreateRow(currentRowIndex);
                    row.CreateCell(0).SetCellValue(keyValuePair.Key);
                    row.CreateCell(1).SetCellValue(keyValuePair.Value);
                    row.GetCell(0).CellStyle = NORMAL_CELL_STYLE;
                    row.GetCell(1).CellStyle = NORMAL_CELL_STYLE;

                    currentRowIndex++;
                }
            }
            // 解析请求的内容
            {
                if (string.IsNullOrEmpty(ReqContentType))
                {
                    IRow bodyRow = sheet.CreateRow(currentRowIndex);
                    bodyRow.CreateCell(0).SetCellValue("解析请求的内容");
                    bodyRow.GetCell(0).CellStyle = BOLD_CELL_STYLE;
                    ICell bodyRowCell = bodyRow.CreateCell(1);
                    bodyRowCell.CellStyle = WRAP_TEXT_CELL_STYLE;

                    bodyRowCell.SetCellValue("无");

                    currentRowIndex++;
                }
                else
                {
                    if (ReqContentType.IndexOf("application/json", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        IRow bodyRow = sheet.CreateRow(currentRowIndex);
                        bodyRow.CreateCell(0).SetCellValue("解析请求的Json");
                        bodyRow.GetCell(0).CellStyle = BOLD_CELL_STYLE;
                        ICell bodyRowCell = bodyRow.CreateCell(1);
                        bodyRowCell.CellStyle = WRAP_TEXT_CELL_STYLE;

                        if (userConfig.IsPrettyPrintJson == true)
                        {
                            JsonData jsonData = JsonMapper.ToObject(ReqJsonParam);
                            JsonWriter jsonWriter = new JsonWriter();
                            jsonWriter.PrettyPrint = true;
                            jsonData.ToJson(jsonWriter);
                            // 这样生成的json字符串换行都是“\r\n”，先全替换为“\n”，然后把开头的“/n”去掉，否则开头多一个空行
                            string text = jsonWriter.TextWriter.ToString().Replace("\r\n", "\n").Remove(0, 1);
                            // 还要将LitJson强制把中文字符变成的unicode码还原回去
                            Regex regex = new Regex(@"(?i)\\[uU]([0-9a-f]{4})");
                            text = regex.Replace(text, delegate (Match m) { return ((char)Convert.ToInt32(m.Groups[1].Value, 16)).ToString(); });

                            if (text.Length > AppConst.MAX_SHOW_TEXT_LENGTH_IN_CELL)
                                bodyRowCell.SetCellValue(string.Concat(AppConst.OVERFLOW_TEXT_TIPS, text.Substring(0, AppConst.MAX_SHOW_TEXT_LENGTH_IN_CELL)));
                            else
                                bodyRowCell.SetCellValue(text);
                        }
                        else
                        {
                            if (ReqJsonParam.Length > AppConst.MAX_SHOW_TEXT_LENGTH_IN_CELL)
                                bodyRowCell.SetCellValue(string.Concat(AppConst.OVERFLOW_TEXT_TIPS, ReqJsonParam.Substring(0, AppConst.MAX_SHOW_TEXT_LENGTH_IN_CELL)));
                            else
                                bodyRowCell.SetCellValue(ReqJsonParam);
                        }

                        currentRowIndex++;
                    }
                    else if (ReqContentType.IndexOf("application/x-www-form-urlencoded", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        IRow titleRow = sheet.CreateRow(currentRowIndex);
                        titleRow.CreateCell(0).SetCellValue("解析请求的表单如下，每行一个键值对");
                        titleRow.GetCell(0).CellStyle = BOLD_CELL_STYLE;

                        currentRowIndex++;

                        foreach (var keyValuePair in ReqFormParamDict)
                        {
                            IRow row = sheet.CreateRow(currentRowIndex);
                            row.CreateCell(0).SetCellValue(keyValuePair.Key);
                            row.CreateCell(1).SetCellValue(keyValuePair.Value);
                            row.GetCell(0).CellStyle = NORMAL_CELL_STYLE;
                            row.GetCell(1).CellStyle = NORMAL_CELL_STYLE;

                            currentRowIndex++;
                        }
                    }
                    else
                        throw new NotImplementedException();
                }
            }

            // 接下来为“服务器响应部分”的标题
            {
                // 标题行设置合并后居中
                IRow titleRow = sheet.CreateRow(currentRowIndex);
                titleRow.HeightInPoints = 22.5f;
                // 标题行水平居中
                ICellStyle titleStyle = workbook.CreateCellStyle();
                titleStyle.BorderTop = BorderStyle.Thin;
                titleStyle.BorderBottom = BorderStyle.Thin;
                titleStyle.BorderLeft = BorderStyle.Thin;
                titleStyle.BorderRight = BorderStyle.Thin;
                titleStyle.VerticalAlignment = VerticalAlignment.Top;
                titleStyle.Alignment = HorizontalAlignment.Center;
                // 标题行的字体
                IFont titleFont = workbook.CreateFont();
                titleFont.FontHeightInPoints = 18;
                titleFont.IsBold = true;
                titleStyle.SetFont(titleFont);
                // 由于NPOI的设计缺陷，要想对合并单元格设置边框，只能先将合并前的单元格进行边框设置后再合并
                for (int columnIndex = 0; columnIndex < 1; columnIndex++)
                {
                    // 这里必须先create出要被合并的单元格，否则NPOI获取不到也就无法设置单元格格式
                    titleRow.CreateCell(columnIndex).CellStyle = titleStyle;
                }

                ICell titleCell = titleRow.GetCell(0);
                titleCell.SetCellValue("服务器响应部分");
                // 合并单元格，四个参数分别是合并起始行index、结束行index、合并起始列index、结束列index
                CellRangeAddress titleRegion = new CellRangeAddress(currentRowIndex, currentRowIndex, 0, 1);
                sheet.AddMergedRegion(titleRegion);
            }
            currentRowIndex++;
            // 状态码
            {
                IRow row = sheet.CreateRow(currentRowIndex);
                row.CreateCell(0).SetCellValue("状态码");
                row.CreateCell(1).SetCellValue(RespStateCode);
                row.GetCell(0).CellStyle = NORMAL_CELL_STYLE;
                row.GetCell(1).CellStyle = NORMAL_CELL_STYLE;
            }
            currentRowIndex++;
            // content-type
            {
                IRow row = sheet.CreateRow(currentRowIndex);
                row.CreateCell(0).SetCellValue(RespHeaderKeyToRealName.ContainsKey("content-type") ? RespHeaderKeyToRealName["content-type"] : "content-type");
                row.CreateCell(1).SetCellValue(string.IsNullOrEmpty(RespContentType) ? "未声明" : RespContentType);
                row.GetCell(0).CellStyle = NORMAL_CELL_STYLE;
                row.GetCell(1).CellStyle = NORMAL_CELL_STYLE;
            }
            currentRowIndex++;
            // 如果有Cookie，输出所有的keyValuePair
            if (RespCookies.Count > 0)
            {
                IRow titleRow = sheet.CreateRow(currentRowIndex);
                string cookieParamName = RespHeaderKeyToRealName.ContainsKey("set-cookie") ? RespHeaderKeyToRealName["set-cookie"] : "set-cookie";
                titleRow.CreateCell(0).SetCellValue($"{cookieParamName}如下，每行一个键值对");
                titleRow.GetCell(0).CellStyle = BOLD_CELL_STYLE;

                currentRowIndex++;
                foreach (Cookie cookie in RespCookies)
                {
                    IRow row = sheet.CreateRow(currentRowIndex);
                    row.CreateCell(0).SetCellValue(cookie.Name);
                    row.GetCell(0).CellStyle = NORMAL_CELL_STYLE;

                    StringBuilder textBuilder = new StringBuilder();
                    textBuilder.Append(cookie.Value).Append("\n");
                    // 当未设置cookie过期时间时，DateTime值为0年1月1日，且用Expired属性判定得到false
                    textBuilder.Append("path=").Append(cookie.Path).Append(";domain=").Append(cookie.Domain).Append(";expires=")
                        .Append(cookie.Expired == false && cookie.Expires == DateTime.MinValue ? "不过期" : cookie.Expires.ToString()).Append(";");
                    if (cookie.HttpOnly)
                        textBuilder.Append("HttpOnly;");
                    if (cookie.Secure)
                        textBuilder.Append("Secure;");
                    row.CreateCell(1).SetCellValue(textBuilder.ToString());
                    row.GetCell(1).CellStyle = WRAP_TEXT_CELL_STYLE;

                    currentRowIndex++;
                }
            }
            // 如果有其他Header字段，输出所有的keyValuePair
            if (RespHeaders.Count > 0)
            {
                IRow titleRow = sheet.CreateRow(currentRowIndex);
                titleRow.CreateCell(0).SetCellValue("其他Header字段，每行一个键值对");
                titleRow.GetCell(0).CellStyle = BOLD_CELL_STYLE;

                currentRowIndex++;
                foreach (var keyValuePair in RespHeaders)
                {
                    IRow row = sheet.CreateRow(currentRowIndex);
                    row.CreateCell(0).SetCellValue(keyValuePair.Key);
                    row.CreateCell(1).SetCellValue(keyValuePair.Value);
                    row.GetCell(0).CellStyle = NORMAL_CELL_STYLE;
                    row.GetCell(1).CellStyle = NORMAL_CELL_STYLE;

                    currentRowIndex++;
                }
            }
            // 解析服务器返回的内容
            {
                if (string.IsNullOrEmpty(RespContentType))
                {
                    IRow bodyRow = sheet.CreateRow(currentRowIndex);
                    bodyRow.CreateCell(0).SetCellValue("解析服务器返回的内容");
                    bodyRow.GetCell(0).CellStyle = BOLD_CELL_STYLE;
                    ICell bodyRowCell = bodyRow.CreateCell(1);
                    bodyRowCell.CellStyle = WRAP_TEXT_CELL_STYLE;

                    bodyRowCell.SetCellValue("无");

                    currentRowIndex++;
                }
                else
                {
                    if (RespContentType.IndexOf("application/json", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        IRow bodyRow = sheet.CreateRow(currentRowIndex);
                        bodyRow.CreateCell(0).SetCellValue("解析服务器返回的Json");
                        bodyRow.GetCell(0).CellStyle = BOLD_CELL_STYLE;
                        ICell bodyRowCell = bodyRow.CreateCell(1);
                        bodyRowCell.CellStyle = WRAP_TEXT_CELL_STYLE;

                        if (userConfig.IsPrettyPrintJson == true)
                        {
                            JsonData jsonData = JsonMapper.ToObject(RespJson);
                            JsonWriter jsonWriter = new JsonWriter();
                            jsonWriter.PrettyPrint = true;
                            jsonData.ToJson(jsonWriter);
                            // 这样生成的json字符串换行都是“\r\n”，先全替换为“\n”，然后把开头的“/n”去掉，否则开头多一个空行
                            string text = jsonWriter.TextWriter.ToString().Replace("\r\n", "\n").Remove(0, 1);
                            // 还要将LitJson强制把中文字符变成的unicode码还原回去
                            Regex regex = new Regex(@"(?i)\\[uU]([0-9a-f]{4})");
                            text = regex.Replace(text, delegate (Match m) { return ((char)Convert.ToInt32(m.Groups[1].Value, 16)).ToString(); });

                            if (text.Length > AppConst.MAX_SHOW_TEXT_LENGTH_IN_CELL)
                                bodyRowCell.SetCellValue(string.Concat(AppConst.OVERFLOW_TEXT_TIPS, text.Substring(0, AppConst.MAX_SHOW_TEXT_LENGTH_IN_CELL)));
                            else
                                bodyRowCell.SetCellValue(text);
                        }
                        else
                        {
                            if (RespJson.Length > AppConst.MAX_SHOW_TEXT_LENGTH_IN_CELL)
                                bodyRowCell.SetCellValue(string.Concat(AppConst.OVERFLOW_TEXT_TIPS, RespJson.Substring(0, AppConst.MAX_SHOW_TEXT_LENGTH_IN_CELL)));
                            else
                                bodyRowCell.SetCellValue(RespJson);
                        }

                        currentRowIndex++;
                    }
                    else if (RespContentType.IndexOf("text/html", StringComparison.CurrentCultureIgnoreCase) != -1 || RespContentType.IndexOf("text/plain", StringComparison.CurrentCultureIgnoreCase) != -1)
                    {
                        IRow bodyRow = sheet.CreateRow(currentRowIndex);
                        bodyRow.CreateCell(0).SetCellValue("解析服务器返回的文本");
                        bodyRow.GetCell(0).CellStyle = BOLD_CELL_STYLE;
                        ICell bodyRowCell = bodyRow.CreateCell(1);
                        bodyRowCell.CellStyle = WRAP_TEXT_CELL_STYLE;

                        if (RespText.Length > AppConst.MAX_SHOW_TEXT_LENGTH_IN_CELL)
                            bodyRowCell.SetCellValue(string.Concat(AppConst.OVERFLOW_TEXT_TIPS, RespText.Substring(0, AppConst.MAX_SHOW_TEXT_LENGTH_IN_CELL)));
                        else
                            bodyRowCell.SetCellValue(RespText);

                        currentRowIndex++;
                    }
                    else
                        throw new NotImplementedException();
                }
            }
        }
    }
}
