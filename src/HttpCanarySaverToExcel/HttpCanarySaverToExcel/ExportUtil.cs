using HttpCanarySaverToExcel.Adapter;
using LitJson;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Web;

namespace HttpCanarySaverToExcel
{
    public class ExportUtil
    {
        /// <summary>
        /// 执行HttpCanary存档转为Excel文件
        /// </summary>
        /// <param name="userConfig">用户配置</param>
        /// <param name="errorString">如果导出失败，返回错误信息</param>
        /// <returns>是否导出成功</returns>
        public static bool ExecuteExport(UserConfig userConfig, out string errorString)
        {
            /**
             * 在选择的HttpCanary存档文件夹中，找到每个记录对应的子文件夹
             */
            int maxRecordSeq = -1;
            bool isInputMaxRecordSeq = (userConfig.IsRevertSaverNumberSeq == true && userConfig.InputMaxRecordSeq > 0);
            if (isInputMaxRecordSeq == true)
                maxRecordSeq = userConfig.InputMaxRecordSeq;
            // key:seq, value:OneRecordFileInfo
            Dictionary<int, OneRecordFileInfo> allRecordFileInfoDict = new Dictionary<int, OneRecordFileInfo>();

            if (userConfig.IsRevertSaverNumberSeq == true)
            {
                // 首次遍历子文件夹，找到或验证最大记录序号
                foreach (string dirPath in Directory.GetDirectories(userConfig.SaverFolderPath))
                {
                    int lastIndexForwardSlash = dirPath.LastIndexOf("\\");
                    string folderName = lastIndexForwardSlash > -1 && lastIndexForwardSlash != dirPath.Length - 1 ? dirPath.Substring(lastIndexForwardSlash + 1) : dirPath;
                    int firstIndexForBracket = folderName.IndexOf("（");
                    string numSeqString = firstIndexForBracket > -1 ? folderName.Substring(0, firstIndexForBracket) : folderName;
                    int numSeq = -1;
                    if (int.TryParse(numSeqString, out numSeq) == false || numSeq < 1)
                    {
                        errorString = $"存档文件夹中含有非法名称的文件名为：{folderName}，请保证存档文件夹下都是由HttpCanary自动命名的子文件夹，如果要指定备注内容，请在序号后面用中文小括号括起来";
                        return false;
                    }
                    // 如果逆序排列，并且指定了最大编号，但出现了大于指定的最大编号的文件夹名，需要报错给用户，因为这样将导致无法与App中顺序一致
                    if (isInputMaxRecordSeq == true && numSeq > maxRecordSeq)
                    {
                        errorString = $"指定了最大记录编号为{maxRecordSeq}，但存档文件夹中存在编号大于它的子文件夹为：{folderName}，程序被迫中止，因为只有在正确输入HttpCanary本次存档的最大编号情况下，逆序后才能与App中显示的序号一致";
                        return false;
                    }
                    // 如果出现相同的序号进行报错
                    if (allRecordFileInfoDict.ContainsKey(numSeq) == true)
                    {
                        errorString = $"存档文件夹中出现了相同的记录编号{numSeq}，程序被迫中止，请保证是HttpCanary正常导出保存的存档文件夹";
                        return false;
                    }
                    if (isInputMaxRecordSeq == false && numSeq > maxRecordSeq)
                        maxRecordSeq = numSeq;
                }
                // 再次遍历，生成记录信息
                foreach (string dirPath in Directory.GetDirectories(userConfig.SaverFolderPath))
                {
                    int lastIndexForwardSlash = dirPath.LastIndexOf("\\");
                    string folderName = lastIndexForwardSlash > -1 && lastIndexForwardSlash != dirPath.Length - 1 ? dirPath.Substring(lastIndexForwardSlash + 1) : dirPath;
                    int firstIndexForBracket = folderName.IndexOf("（");
                    string numSeqString = firstIndexForBracket > -1 ? folderName.Substring(0, firstIndexForBracket) : folderName;
                    int numSeq = int.Parse(numSeqString);
                    // 逆序后的序号
                    numSeq = maxRecordSeq + 1 - numSeq;

                    string remark = null;
                    if (firstIndexForBracket > -1)
                    {
                        int lastIndexForBracket = folderName.LastIndexOf("）");
                        if (lastIndexForBracket == -1)
                        {
                            errorString = $"存档文件夹中含有非法名称的文件名为：{folderName}，其含有左括号，但没有右括号，若要指定备注内容，请在序号后面用中文小括号将备注内容括起来";
                            return false;
                        }
                        else
                            remark = folderName.Substring(firstIndexForBracket + 1, lastIndexForBracket - firstIndexForBracket - 1);
                    }

                    OneRecordFileInfo info = new OneRecordFileInfo();
                    info.NumberSeq = numSeq;
                    info.FullFolderPath = dirPath;
                    info.Remark = remark;
                    allRecordFileInfoDict.Add(numSeq, info);
                }
            }
            else
            {
                foreach (string dirPath in Directory.GetDirectories(userConfig.SaverFolderPath))
                {
                    int lastIndexForwardSlash = dirPath.LastIndexOf("\\");
                    string folderName = lastIndexForwardSlash > -1 && lastIndexForwardSlash != dirPath.Length - 1 ? dirPath.Substring(lastIndexForwardSlash + 1) : dirPath;
                    int firstIndexForBracket = folderName.IndexOf("（");
                    string numSeqString = firstIndexForBracket > -1 ? folderName.Substring(0, firstIndexForBracket) : folderName;
                    int numSeq = -1;
                    if (int.TryParse(numSeqString, out numSeq) == false || numSeq < 1)
                    {
                        errorString = $"存档文件夹中含有非法名称的文件名为：{folderName}，请保证存档文件夹下都是由HttpCanary自动命名的子文件夹，如果要指定备注内容，请在序号后面用中文小括号括起来";
                        return false;
                    }
                    // 如果出现相同的序号进行报错
                    if (allRecordFileInfoDict.ContainsKey(numSeq) == true)
                    {
                        errorString = $"存档文件夹中出现了相同的记录编号{numSeq}，程序被迫中止，请保证是HttpCanary正常导出保存的存档文件夹";
                        return false;
                    }
                    if (numSeq > maxRecordSeq)
                        maxRecordSeq = numSeq;

                    string remark = null;
                    if (firstIndexForBracket > -1)
                    {
                        int lastIndexForBracket = folderName.LastIndexOf("）");
                        if (lastIndexForBracket == -1)
                        {
                            errorString = $"存档文件夹中含有非法名称的文件名为：{folderName}，其含有左括号，但没有右括号，若要指定备注内容，请在序号后面用中文小括号将备注内容括起来";
                            return false;
                        }
                        else
                            remark = folderName.Substring(firstIndexForBracket + 1, lastIndexForBracket - firstIndexForBracket - 1);
                    }

                    OneRecordFileInfo info = new OneRecordFileInfo();
                    info.NumberSeq = numSeq;
                    info.FullFolderPath = dirPath;
                    info.Remark = remark;
                    allRecordFileInfoDict.Add(numSeq, info);
                }
            }
            if (allRecordFileInfoDict.Count == 0)
            {
                errorString = $"存档文件夹中未找到任何记录对应的子文件夹，请重新选择正确的存档目录";
                return false;
            }

            /**
             * 解析每条记录
             */
            // 存储所有的OneOverviewRecord
            List<OneOverviewRecord> allOverviewRecord = new List<OneOverviewRecord>();
            // 存储所有记录对应的Adapter
            List<AdapterBase> allAdapter = new List<AdapterBase>();

            for (int i = 1; i <= maxRecordSeq; i++)
            {
                if (allRecordFileInfoDict.ContainsKey(i) == false)
                    continue;

                AdapterBase adapter = null;
                string analyzeErrorString = null;
                try
                {
                    OneOverviewRecord overviewRecord = _AnalyzeOneRecord(userConfig, allRecordFileInfoDict[i], out adapter, out analyzeErrorString);
                    if (analyzeErrorString != null)
                    {
                        errorString = $"解析“{allRecordFileInfoDict[i].FullFolderPath}”中的记录失败，原因为：{analyzeErrorString}，程序被迫中止";
                        return false;
                    }
                    else
                    {
                        allOverviewRecord.Add(overviewRecord);
                        allAdapter.Add(adapter);
                    }
                }
                catch (Exception e)
                {
                    errorString = $"解析“{allRecordFileInfoDict[i].FullFolderPath}”中的记录抛出异常，程序被迫中止，错误堆栈信息为：\n{e.ToString()}";
                    return false;
                }
            }

            /**
             * 最终生成Excel文件
             */
            try
            {
                _WriteExcel(userConfig, allOverviewRecord, allAdapter);
                errorString = null;
                return true;
            }
            catch (Exception e)
            {
                errorString = $"最终生成Excel文件抛出异常，错误堆栈信息为：\n{e.ToString()}";
                return false;
            }
        }

        private static void _WriteExcel(UserConfig userConfig, List<OneOverviewRecord> allOverviewRecord, List<AdapterBase> allAdapter)
        {
            using (var fs = new FileStream(userConfig.ExportExcelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
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

                /**
                 * 总览Sheet
                 */
                ISheet overviewSheet = workbook.CreateSheet(AppConst.OVERVIEW_SHEET_NAME);

                // 标题行设置合并后居中
                IRow titleRow = overviewSheet.CreateRow(0);
                titleRow.HeightInPoints = 46.5f;
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
                titleFont.FontHeightInPoints = 36;
                titleFont.IsBold = true;
                titleStyle.SetFont(titleFont);
                // 由于NPOI的设计缺陷，要想对合并单元格设置边框，只能先将合并前的单元格进行边框设置后再合并
                for (int columnIndex = 0; columnIndex < AppConst.OVERVIEW_COLUMN_TITLES.Length; columnIndex++)
                {
                    // 这里必须先create出要被合并的单元格，否则NPOI获取不到也就无法设置单元格格式
                    titleRow.CreateCell(columnIndex).CellStyle = titleStyle;
                }

                ICell titleCell = titleRow.GetCell(0);
                titleCell.SetCellValue(AppConst.OVERVIEW_SHEET_NAME);
                // 合并单元格，四个参数分别是合并起始行index、结束行index、合并起始列index、结束列index
                CellRangeAddress titleRegion = new CellRangeAddress(0, 0, 0, AppConst.OVERVIEW_COLUMN_TITLES.Length - 1);
                overviewSheet.AddMergedRegion(titleRegion);

                // 第2行是说明行
                StringBuilder noticeSB = new StringBuilder();
                noticeSB.Append("注意：\n");
                noticeSB.Append(userConfig.IsRevertSaverNumberSeq == true ? AppConst.IS_REVERT_SAVER_NUMBER_SEQ_NOTICE_TEXT : AppConst.NOT_REVERT_SAVER_NUMBER_SEQ_NOTICE_TEXT);
                IRow noticeRow = overviewSheet.CreateRow(1);
                noticeRow.HeightInPoints = 30f;
                // 由于NPOI的设计缺陷，要想对合并单元格设置边框，只能先将合并前的单元格进行边框设置后再合并
                for (int columnIndex = 0; columnIndex < AppConst.OVERVIEW_COLUMN_TITLES.Length; columnIndex++)
                {
                    // 这里必须先create出要被合并的单元格，否则NPOI获取不到也就无法设置单元格格式
                    noticeRow.CreateCell(columnIndex).CellStyle = WRAP_TEXT_CELL_STYLE;
                }
                noticeRow.GetCell(0).SetCellValue(noticeSB.ToString());
                CellRangeAddress noticeRegion = new CellRangeAddress(1, 1, 0, AppConst.OVERVIEW_COLUMN_TITLES.Length - 1);
                overviewSheet.AddMergedRegion(noticeRegion);

                // 第3行写入各字段列的列名
                IRow columnNameRow = overviewSheet.CreateRow(2);
                for (int columnIndex = 0; columnIndex < AppConst.OVERVIEW_COLUMN_TITLES.Length; columnIndex++)
                {
                    ICell cell = columnNameRow.CreateCell(columnIndex);
                    cell.SetCellValue(AppConst.OVERVIEW_COLUMN_TITLES[columnIndex]);
                    cell.CellStyle = NORMAL_CELL_STYLE;
                }

                // 从第4行开始写入记录的总览内容
                for (int index = 0; index < allOverviewRecord.Count; index++)
                {
                    int rowIndex = index + AppConst.START_RECORD_ROW_INDEX;
                    IRow row = overviewSheet.CreateRow(rowIndex);
                    OneOverviewRecord overviewRecord = allOverviewRecord[index];
                    // 序号
                    row.CreateCell(0).SetCellValue(overviewRecord.RecordFileInfo.NumberSeq);
                    // App包名
                    row.CreateCell(1).SetCellValue(overviewRecord.AppPackageName);
                    // 协议类型
                    row.CreateCell(2).SetCellValue(overviewRecord.NetType.ToString());
                    // URL
                    row.CreateCell(3).SetCellValue(overviewRecord.Url);
                    // 跳转到详情Sheet（如果有详情页，添加跳转到对应Sheet的超链接，否则打印没有详情页的原因）
                    if (overviewRecord.IsGenerateDetialSheet == true)
                    {
                        ICell detialCell = row.CreateCell(4);
                        detialCell.SetCellValue("点击跳转到对应详情Sheet表");
                        // 设置超链接到对应Sheet表
                        XSSFHyperlink hyperlink = new XSSFHyperlink(HyperlinkType.Document);
                        hyperlink.Address = $"'{overviewRecord.RecordFileInfo.NumberSeq}'!A1";
                        detialCell.Hyperlink = hyperlink;
                        detialCell.CellStyle = HYPERLINK_CELL_STYLE;
                    }
                    else
                    {
                        ICell cell = row.CreateCell(4);
                        cell.SetCellValue(overviewRecord.NotGenerateDetialSheetReason);
                        cell.CellStyle = NORMAL_CELL_STYLE;
                    }
                    // 备注
                    row.CreateCell(5).SetCellValue(overviewRecord.RecordFileInfo.Remark);
                }
                // 调整列宽
                for (int columnIndex = 0; columnIndex < AppConst.OVERVIEW_COLUMN_TITLES.Length; columnIndex++)
                {
                    // 第2个参数的单位是1/256个字符宽度
                    overviewSheet.SetColumnWidth(columnIndex, (int)(AppConst.OVERVIEW_SHEET_COLUMN_WIDTH[columnIndex] * 256f));
                }
                // 设置总览内容单元格的格式
                for (int rowIndex = AppConst.START_RECORD_ROW_INDEX; rowIndex < AppConst.START_RECORD_ROW_INDEX + allOverviewRecord.Count; rowIndex++)
                {
                    for (int columnIndex = 0; columnIndex < AppConst.OVERVIEW_COLUMN_TITLES.Length; columnIndex++)
                    {
                        // 排除掉“跳转到详情Sheet”这一列，因为已经设置过单元格格式
                        if (columnIndex == 4)
                            continue;

                        ICell cell = overviewSheet.GetRow(rowIndex).GetCell(columnIndex);
                        cell.CellStyle = NORMAL_CELL_STYLE;
                    }
                }

                /**
                 * 写入各个详情Sheet表
                 */
                for (int i = 0; i < allOverviewRecord.Count; i++)
                {
                    if (allOverviewRecord[i].IsGenerateDetialSheet == true)
                        allAdapter[i].WriteExcelSheet(workbook, userConfig, allOverviewRecord[i]);
                }

                /**
                 * 最终写入Excel文件
                 */
                workbook.Write(fs);
            }
        }

        /// <summary>
        /// 解析一个记录文件夹
        /// </summary>
        /// <param name="userConfig"></param>
        /// <param name="recordFileInfo"></param>
        /// <param name="adapter">out返回解析的记录对应的Adapter</param>
        /// <param name="errorString">out返回发生错误后的详细错误原因</param>
        /// <returns>如果该记录要被筛选掉或解析出错，返回null，否则正常返回该记录对应在总览Sheet中展示的信息</returns>
        private static OneOverviewRecord _AnalyzeOneRecord(UserConfig userConfig, OneRecordFileInfo recordFileInfo, out AdapterBase adapter, out string errorString)
        {
            OneOverviewRecord overviewRecord = new OneOverviewRecord();
            overviewRecord.RecordFileInfo = recordFileInfo;

            /**
             * 获取存档子文件夹下的所有文件名
             */
            // key:fileNameWithExtension, value:FileInfo
            Dictionary<string, FileInfo> allFileInfoDict = new Dictionary<string, FileInfo>();
            DirectoryInfo directoryInfo = new DirectoryInfo(recordFileInfo.FullFolderPath);
            FileInfo[] allFileInfo = directoryInfo.GetFiles();
            foreach (FileInfo fileInfo in allFileInfo)
            {
                allFileInfoDict.Add(fileInfo.Name, fileInfo);
            }
            /**
             * 根据存档内文件名判断通讯协议类型
             */
            // TCP
            if (allFileInfoDict.ContainsKey("tcp.json"))
            {
                overviewRecord.NetType = NetTypeEnum.TCP;
                TcpAdapter tcpAdapter = new TcpAdapter();
                adapter = tcpAdapter;

                string tcpInfoJson = File.ReadAllText(allFileInfoDict["tcp.json"].FullName, Encoding.UTF8);
                JsonData jsonData = JsonMapper.ToObject(tcpInfoJson);
                // 这里要处理疑似HttpCanary的bug，有些请求的request.json文件中确实赋值为null
                tcpAdapter.AppPackageName = (jsonData["app"] != null ? jsonData["app"].ToString() : "null");
                if (userConfig.TargetAppPackageNames != null && userConfig.TargetAppPackageNames.Contains(tcpAdapter.AppPackageName) == false)
                {
                    errorString = null;
                    return null;
                }

                tcpAdapter.RemoteIpAndPort = string.Concat(jsonData["remoteIp"].ToString(), ":", jsonData["remotePort"].ToString());
                overviewRecord.AppPackageName = tcpAdapter.AppPackageName;
                overviewRecord.Url = tcpAdapter.RemoteIpAndPort;
                overviewRecord.IsGenerateDetialSheet = false;
                overviewRecord.NotGenerateDetialSheetReason = "暂不支持解析TCP";
                errorString = null;
                return overviewRecord;
            }
            // UDP
            else if (allFileInfoDict.ContainsKey("udp.json"))
            {
                overviewRecord.NetType = NetTypeEnum.UDP;
                UdpAdapter udpAdapter = new UdpAdapter();
                adapter = udpAdapter;

                string tcpInfoJson = File.ReadAllText(allFileInfoDict["udp.json"].FullName, Encoding.UTF8);
                JsonData jsonData = JsonMapper.ToObject(tcpInfoJson);
                // 这里要处理疑似HttpCanary的bug，有些请求的request.json文件中确实赋值为null
                udpAdapter.AppPackageName = (jsonData["app"] != null ? jsonData["app"].ToString() : "null");
                if (userConfig.TargetAppPackageNames != null && userConfig.TargetAppPackageNames.Contains(udpAdapter.AppPackageName) == false)
                {
                    errorString = null;
                    return null;
                }
                udpAdapter.RemoteIpAndPort = string.Concat(jsonData["remoteIp"].ToString(), ":", jsonData["remotePort"].ToString());
                overviewRecord.AppPackageName = udpAdapter.AppPackageName;
                overviewRecord.Url = udpAdapter.RemoteIpAndPort;
                overviewRecord.IsGenerateDetialSheet = false;
                overviewRecord.NotGenerateDetialSheetReason = "暂不支持解析UDP";
                errorString = null;
                return overviewRecord;
            }
            // HTTP和HTTPS需要具体分析url进行区分
            else if (allFileInfoDict.ContainsKey("request.json"))
            {
                HttpAdapter httpAdapter = new HttpAdapter();
                adapter = httpAdapter;

                /**
                 * request.json文件记录请求的总览信息
                 */
                string requestInfoJson = File.ReadAllText(allFileInfoDict["request.json"].FullName, Encoding.UTF8);
                JsonData reqJsonData = JsonMapper.ToObject(requestInfoJson);
                // 这里要处理疑似HttpCanary的bug，有些请求的request.json文件中确实赋值为null
                httpAdapter.AppPackageName = (reqJsonData["app"] != null ? reqJsonData["app"].ToString() : "null");
                if (userConfig.TargetAppPackageNames != null && userConfig.TargetAppPackageNames.Contains(httpAdapter.AppPackageName) == false)
                {
                    errorString = null;
                    return null;
                }
                httpAdapter.RemoteIpAndPort = string.Concat(reqJsonData["remoteIp"].ToString(), ":", reqJsonData["remotePort"].ToString());

                string url = reqJsonData["url"].ToString();
                overviewRecord.AppPackageName = httpAdapter.AppPackageName;
                overviewRecord.Url = url;
                if (url.StartsWith("https:"))
                    overviewRecord.NetType = NetTypeEnum.HTTPS;
                else
                    overviewRecord.NetType = NetTypeEnum.HTTP;

                // 处理URL解码
                if (userConfig.IsUrlDecode == true)
                    url = HttpUtility.UrlDecode(url);

                // 处理queryString
                int questionMarkIndex = url.IndexOf("?");
                if (questionMarkIndex == -1)
                    httpAdapter.ReqtUrl = url;
                else
                {
                    httpAdapter.ReqtUrl = url.Substring(0, questionMarkIndex);
                    string queryString = url.Substring(questionMarkIndex + 1);
                    string[] allKeyAndValuePair = queryString.Split(new string[] { "&" }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string oneKeyAndValue in allKeyAndValuePair)
                    {
                        string[] keyAndValue = oneKeyAndValue.Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries);
                        // 这里要处理value为空的情况
                        if (keyAndValue.Length == 2)
                            httpAdapter.ReqQueryStringDict.Add(keyAndValue[0], keyAndValue[1]);
                        else if (keyAndValue.Length == 1)
                            httpAdapter.ReqQueryStringDict.Add(keyAndValue[0], string.Empty);
                        else
                        {
                            errorString = $"解析请求中的QueryString发现错误的键值对为：{oneKeyAndValue}";
                            return null;
                        }
                    }
                }

                httpAdapter.ReqMethod = reqJsonData["method"].ToString();
                // 解析headers
                if (reqJsonData.ContainsKey("headers"))
                {
                    JsonData headers = reqJsonData["headers"];
                    foreach (string key in headers.Keys)
                    {
                        // 用户设置要忽略掉的抛弃
                        if (userConfig.IgnoreReqHeaderName.Contains(key.ToLower()))
                            continue;

                        // 处理一些固定分析的字段
                        if ("content-type".Equals(key, StringComparison.CurrentCultureIgnoreCase))
                        {
                            httpAdapter.ReqHeaderKeyToRealName.Add("content-type", key);
                            httpAdapter.ReqContentType = headers[key].ToString();
                        }
                        else if ("cookie".Equals(key, StringComparison.CurrentCultureIgnoreCase))
                        {
                            httpAdapter.ReqHeaderKeyToRealName.Add("cookie", key);
                            string cookieString = headers[key].ToString();
                            // cookie中不同键值对通过英文分号隔开，但注意分号后还有空格
                            string[] cookies = cookieString.Split(new string[] { ";" }, System.StringSplitOptions.RemoveEmptyEntries);
                            foreach (string oneCookieKeyValuePair in cookies)
                            {
                                string[] keyAndValue = oneCookieKeyValuePair.Trim().Split(new string[] { "=" }, System.StringSplitOptions.RemoveEmptyEntries);
                                if (keyAndValue.Length != 2)
                                {
                                    errorString = $"解析请求中的cookie发现键值对错误，出错的cookie为：{oneCookieKeyValuePair}";
                                    return null;
                                }
                                else
                                {
                                    string cookieKey = keyAndValue[0].Trim();
                                    string cookieValue = keyAndValue[1].Trim();
                                    Cookie cookie = new Cookie();
                                    cookie.Name = cookieKey;
                                    cookie.Value = cookieValue;
                                    httpAdapter.ReqCookies.Add(cookie);
                                }
                            }
                        }
                        // 其余字段进行记录
                        else
                        {
                            string value = headers[key].ToString();
                            httpAdapter.ReqHeaders.Add(key, value);
                        }
                    }
                }
                /**
                 * 根据content-type判断HttpCancry保存的请求文件类型
                 */
                // 无content-type
                if (httpAdapter.ReqContentType == null)
                {
                    // 如果header中没有content-type，按标准应该是没有请求内容，但如果在存档文件夹中找到request开头的文件，但不是request.hcy和request.json
                    // 说明请求不规范，而用程序无法根据规定的规则判断出该以哪种content-type进行处理，只能报错给用户
                    foreach (string fileName in allFileInfoDict.Keys)
                    {
                        if (fileName.StartsWith("request") && "request.hcy".Equals(fileName) == false && "request.json".Equals(fileName) == false)
                        {
                            errorString = $"未在request.json文件的header信息中找到content-type，但在存档文件夹中找到名为{fileName}的文件，没有content-type无法解析";
                            return null;
                        }
                    }
                }
                // application/json
                else if (httpAdapter.ReqContentType.IndexOf("application/json", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    if (allFileInfoDict.ContainsKey("request_body.json"))
                    {
                        string requestJson = File.ReadAllText(allFileInfoDict["request_body.json"].FullName, Encoding.UTF8);
                        httpAdapter.ReqJsonParam = requestJson;
                    }
                    else
                    {
                        errorString = "解析header得到content-type为application/json，但找不到对应的request_body.json文件";
                        return null;
                    }
                }
                // application/x-www-form-urlencoded
                else if (httpAdapter.ReqContentType.IndexOf("application/x-www-form-urlencoded", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    if (allFileInfoDict.ContainsKey("request_body.bin"))
                    {
                        // 注意是经过url编码的
                        string body = File.ReadAllText(allFileInfoDict["request_body.bin"].FullName, Encoding.UTF8);
                        if (userConfig.IsUrlDecode == true)
                            body = HttpUtility.UrlDecode(body);

                        string[] allKeyAndValuePair = body.Split(new string[] { "&" }, StringSplitOptions.RemoveEmptyEntries);
                        httpAdapter.ReqFormParamDict = new Dictionary<string, string>();
                        foreach (string oneKeyAndValue in allKeyAndValuePair)
                        {
                            string[] keyAndValue = oneKeyAndValue.Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries);
                            httpAdapter.ReqFormParamDict.Add(keyAndValue[0], keyAndValue[1]);
                        }
                    }
                    else
                    {
                        errorString = "解析header得到content-type为application/json，但找不到对应的request_body.json文件";
                        return null;
                    }
                }
                else
                {
                    overviewRecord.IsGenerateDetialSheet = false;
                    overviewRecord.NotGenerateDetialSheetReason = $"暂未支持content-type为{httpAdapter.ReqContentType}的HTTP请求";
                    errorString = null;
                    return overviewRecord;
                }

                /**
                 * response.json文件记录服务器响应的总览信息
                 */
                if (allFileInfoDict.ContainsKey("response.json") == false)
                {
                    errorString = "在该HTTP请求存档中，无法找到服务器响应的总览信息文件response.json";
                    return null;
                }
                string responseInfoJson = File.ReadAllText(allFileInfoDict["response.json"].FullName, Encoding.UTF8);
                JsonData responseJsonData = JsonMapper.ToObject(responseInfoJson);
                httpAdapter.RespStateCode = int.Parse(responseJsonData["code"].ToString());
                // 解析headers
                if (responseJsonData.ContainsKey("headers"))
                {
                    JsonData headers = responseJsonData["headers"];
                    foreach (string key in headers.Keys)
                    {
                        // 用户设置要忽略掉的抛弃
                        if (userConfig.IgnoreRespHeaderName.Contains(key.ToLower()))
                            continue;

                        // 处理一些固定分析的字段
                        if ("content-type".Equals(key, StringComparison.CurrentCultureIgnoreCase))
                        {
                            httpAdapter.RespHeaderKeyToRealName.Add("content-type", key);
                            httpAdapter.RespContentType = headers[key].ToString();
                        }
                        // 注意可能设置多个cookie
                        else if ("set-cookie".Equals(key, StringComparison.CurrentCultureIgnoreCase))
                        {
                            // 由于json本身的格式限制，当设置多个cookie时，因为key名均为“set-cookie”，会导致json无法全部存储
                            // 所以只能从HttpCanary生成的response.hcy中取得所有的cookie
                            httpAdapter.RespHeaderKeyToRealName.Add("set-cookie", key);

                            if (allFileInfoDict.ContainsKey("response.hcy") == false)
                            {
                                errorString = "在response.json中发现cookie字段，但存档中没有response.hcy";
                                return null;
                            }
                            string[] textAllLines = File.ReadAllLines(allFileInfoDict["response.hcy"].FullName, Encoding.UTF8);
                            foreach (string line in textAllLines)
                            {
                                if (line.StartsWith("set-cookie:", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    string cookieString = line.Remove(0, "set-cookie:".Length);
                                    Cookie cookie = new Cookie();
                                    // 一个cookie中通过英文分号分隔不同参数
                                    string[] cookieParams = cookieString.Split(new string[] { ";" }, System.StringSplitOptions.RemoveEmptyEntries);
                                    foreach (string oneParam in cookieParams)
                                    {
                                        string paramTrim = oneParam.Trim();
                                        if ("secure".Equals(paramTrim, StringComparison.CurrentCultureIgnoreCase))
                                            cookie.Secure = true;
                                        else if ("httponly".Equals(paramTrim, StringComparison.CurrentCultureIgnoreCase))
                                            cookie.HttpOnly = true;
                                        else
                                        {
                                            string[] keyAndValue = paramTrim.Split(new string[] { "=" }, System.StringSplitOptions.RemoveEmptyEntries);
                                            if (keyAndValue.Length != 2)
                                            {
                                                errorString = $"解析cookie中的参数错误，当前cookie参数为：{paramTrim}";
                                                return null;
                                            }
                                            cookie.Name = keyAndValue[0].Trim();
                                            cookie.Value = keyAndValue[1].Trim();
                                            // Cookie类中没有存储max-age的字段，只能跳过不记录
                                            if ("max-age".Equals(key, StringComparison.CurrentCultureIgnoreCase))
                                                continue;
                                            else if ("domain".Equals(key, StringComparison.CurrentCultureIgnoreCase))
                                                cookie.Domain = cookie.Value;
                                            else if ("path".Equals(key, StringComparison.CurrentCultureIgnoreCase))
                                                cookie.Path = cookie.Value;
                                            else if ("expires".Equals(key, StringComparison.CurrentCultureIgnoreCase))
                                            {
                                                // HttpCanary中expires的时间格式形如“Sun, 01-Jan-2022 00:00:00 GMT”
                                                string timeString = cookie.Value.Replace("GMT", "+00");
                                                DateTime expiresDateTime = DateTime.ParseExact(timeString, AppConst.GMT_TIME_FORMAT, AppConst.EN_CULTURE_INFO);
                                                cookie.Expires = expiresDateTime;
                                            }
                                        }
                                    }
                                    httpAdapter.RespCookies.Add(cookie);
                                }
                            }
                        }
                        // 其余字段进行记录
                        else
                        {
                            string value = headers[key].ToString();
                            httpAdapter.RespHeaders.Add(key, value);
                        }
                    }
                }
                /**
                 * 根据content-type判断HttpCancry保存的服务器响应文件类型
                 */
                // 无content-type
                if (httpAdapter.RespContentType == null)
                {
                    // 如果header中没有content-type，按标准应该是没有服务器响应内容，但如果在存档文件夹中找到response开头的文件，但不是response.hcy和response.json
                    // 说明响应不规范，而用程序无法根据规定的规则判断出该以哪种content-type进行处理，只能报错给用户
                    foreach (string fileName in allFileInfoDict.Keys)
                    {
                        if (fileName.StartsWith("response") && "response.hcy".Equals(fileName) == false && "response.json".Equals(fileName) == false)
                        {
                            errorString = $"未在response.json文件的header信息中找到content-type，但在存档文件夹中找到名为{fileName}的文件，没有content-type无法解析";
                            return null;
                        }
                    }
                }
                // application/json
                else if (httpAdapter.RespContentType.IndexOf("application/json", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    if (allFileInfoDict.ContainsKey("response_body.json"))
                    {
                        string responseJson = File.ReadAllText(allFileInfoDict["response_body.json"].FullName, Encoding.UTF8);
                        httpAdapter.RespJson = responseJson;
                    }
                    else
                    {
                        errorString = "解析header得到content-type为application/json，但找不到对应的response_body.json文件";
                        return null;
                    }
                }
                // text/html
                else if (httpAdapter.RespContentType.IndexOf("text/html", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    if (allFileInfoDict.ContainsKey("response_body.html"))
                    {
                        string htmlString = File.ReadAllText(allFileInfoDict["response_body.html"].FullName, Encoding.UTF8);
                        httpAdapter.RespText = htmlString;
                    }
                    else
                    {
                        errorString = "解析header得到content-type为text/html，但找不到对应的response_body.html文件";
                        return null;
                    }
                }
                // text/plain
                else if (httpAdapter.RespContentType.IndexOf("text/plain", StringComparison.CurrentCultureIgnoreCase) != -1)
                {
                    if (allFileInfoDict.ContainsKey("response_body.txt"))
                    {
                        string text = File.ReadAllText(allFileInfoDict["response_body.txt"].FullName, Encoding.UTF8);
                        httpAdapter.RespText = text;
                    }
                    else
                    {
                        errorString = "解析header得到content-type为text/plain，但找不到对应的response_body.txt文件";
                        return null;
                    }
                }
                else
                {
                    overviewRecord.IsGenerateDetialSheet = false;
                    overviewRecord.NotGenerateDetialSheetReason = $"暂未支持content-type为{httpAdapter.RespContentType}的HTTP响应";
                    errorString = null;
                    return overviewRecord;
                }

                overviewRecord.IsGenerateDetialSheet = true;
                errorString = null;
                return overviewRecord;
            }
            else
            {
                overviewRecord.NetType = NetTypeEnum.UNIMPLEMENTED;
                UnimplementedAdapter unimplementedAdapter = new UnimplementedAdapter();
                adapter = unimplementedAdapter;
                overviewRecord.IsGenerateDetialSheet = false;
                overviewRecord.NotGenerateDetialSheetReason = "暂不支持解析此通讯协议";
                errorString = null;
                return overviewRecord;
            }
        }
    }

    /// <summary>
    /// 用于记录一条记录对应的文件夹信息
    /// </summary>
    public class OneRecordFileInfo
    {
        // 序号
        public int NumberSeq { get; set; }
        // 完整的文件夹路径
        public string FullFolderPath { get; set; }
        // 备注（文件名中在数字序号之后的中文小括号里的内容为备注）
        public string Remark { get; set; }
    }
}
