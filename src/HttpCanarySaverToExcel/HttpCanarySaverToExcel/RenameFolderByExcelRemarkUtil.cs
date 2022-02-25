using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;

namespace HttpCanarySaverToExcel
{
    public class AnalyzeRenameFolderByExcelRemarkInfo
    {
        // key:seq, value:当前旧的存档子文件夹路径
        public Dictionary<int, string> SeqToOldFolderName { get; set; } = new Dictionary<int, string>();
        // key:seq（原本应对应的存档文件夹的序号，在逆序情况下，不同于Excel文件中的编号）, value:新的备注名称
        public Dictionary<int, string> SeqToNewFolderName { get; set; } = new Dictionary<int, string>();
        // 是否逆序存档编号
        public bool IsRevertSaverNumberSeq { get; set; }
        // 如果是逆序，返回最大记录序号（由用户指定或者自动遍历统计出来）
        public int MaxRecordSeq { get; set; }
    }

    public class RenameFolderByExcelRemarkUtil
    {
        /// <summary>
        /// 分析Excel文件和存档文件夹，得出要进行重命名的具体情况
        /// </summary>
        /// <param name="excelFilePath">用本工具生成的Excel文件路径</param>
        /// <param name="saverFolderPath">Excel文件对应的HttpCanary存档文件夹路径</param>
        /// <param name="isRevertSaverNumberSeq">是否将存档编号进行逆序，以便和HttpCanary中显示的顺序保持一致</param>
        /// <param name="inputMaxRecordSeq">如果选择了进行逆序，可以指定最大的纪录编号，为0表示不指定</param>
        /// <param name="errorString">如果执行失败，返回错误原因</param>
        /// <returns>分析的结果信息</returns>
        public static AnalyzeRenameFolderByExcelRemarkInfo AnalyzeRenameFolderByExcelRemark(string excelFilePath, string saverFolderPath, bool isRevertSaverNumberSeq, int inputMaxRecordSeq, out string errorString)
        {
            AnalyzeRenameFolderByExcelRemarkInfo info = new AnalyzeRenameFolderByExcelRemarkInfo();

            int maxRecordSeq = -1;
            bool isInputMaxRecordSeq = (isRevertSaverNumberSeq == true && inputMaxRecordSeq > 0);
            if (isInputMaxRecordSeq == true)
                maxRecordSeq = inputMaxRecordSeq;
            /**
             * 在选择的HttpCanary存档文件夹中，找到并记录每条记录对应子文件夹的名称
             */
            foreach (string dirPath in Directory.GetDirectories(saverFolderPath))
            {
                int lastIndexForwardSlash = dirPath.LastIndexOf("\\");
                string folderName = lastIndexForwardSlash > -1 && lastIndexForwardSlash != dirPath.Length - 1 ? dirPath.Substring(lastIndexForwardSlash + 1) : dirPath;
                int firstIndexForBracket = folderName.IndexOf("（");
                string numSeqString = firstIndexForBracket > -1 ? folderName.Substring(0, firstIndexForBracket) : folderName;
                int numSeq = -1;
                if (int.TryParse(numSeqString, out numSeq) == false || numSeq < 1)
                {
                    errorString = $"存档文件夹中含有非法名称的文件名为：{folderName}，请保证存档文件夹下都是由HttpCanary自动命名的子文件夹，如果要指定备注内容，请在序号后面用中文小括号括起来";
                    return null;
                }
                // 如果逆序排列，并且指定了最大编号，但出现了大于指定的最大编号的文件夹名，需要报错给用户，因为这样将导致无法与App中顺序一致
                if (isInputMaxRecordSeq == true && numSeq > maxRecordSeq)
                {
                    errorString = $"指定了最大记录编号为{maxRecordSeq}，但存档文件夹中存在编号大于它的子文件夹为：{folderName}，程序被迫中止，因为只有在正确输入HttpCanary本次存档的最大编号情况下，逆序后才能与App中显示的序号一致";
                    return null;
                }
                // 如果出现相同的序号进行报错
                if (info.SeqToOldFolderName.ContainsKey(numSeq) == true)
                {
                    errorString = $"存档文件夹中出现了相同的记录编号{numSeq}，程序被迫中止，请保证是HttpCanary正常导出保存的存档文件夹";
                    return null;
                }

                info.SeqToOldFolderName.Add(numSeq, dirPath);
                if (isInputMaxRecordSeq == false && numSeq > maxRecordSeq)
                    maxRecordSeq = numSeq;
            }
            /**
             * 从Excel文件中读取每条记录的新备注名称
             */
            using (var fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fs);
                ISheet overviewSheet = workbook.GetSheet(AppConst.OVERVIEW_SHEET_NAME);

                // 先分析判断Excel表中记录的是否逆序与用户选择是否一致
                bool isRevertInExcel;
                string noticeText = overviewSheet.GetRow(1).GetCell(0).ToString().Trim();
                if (noticeText.Contains(AppConst.IS_REVERT_SAVER_NUMBER_SEQ_NOTICE_TEXT))
                    isRevertInExcel = true;
                else if (noticeText.Contains(AppConst.NOT_REVERT_SAVER_NUMBER_SEQ_NOTICE_TEXT))
                    isRevertInExcel = false;
                else
                {
                    errorString = $"从Excel文件{AppConst.OVERVIEW_SHEET_NAME}表中的注意说明中，无法判断当初生成时是否选择逆序存档编号，请不要修改本工具生成的Excel文件格式，本程序无法执行";
                    return null;
                }
                if (isRevertInExcel != isRevertSaverNumberSeq)
                {
                    errorString = $"从Excel文件{AppConst.OVERVIEW_SHEET_NAME}表中的注意说明中，读取到本文件{(isRevertInExcel ? "进行了逆序存档编号" : "没有进行逆序存档编号")}，而您选择的与之相反，请重新确定是否选择正确";
                    return null;
                }

                // Excel表中某行记录的序号
                int seq;
                for (int rowIndex = AppConst.START_RECORD_ROW_INDEX; rowIndex <= overviewSheet.LastRowNum; rowIndex++)
                {
                    IRow row = overviewSheet.GetRow(rowIndex);
                    string seqCellStr = row.GetCell(0).ToString().Trim();
                    if (string.IsNullOrEmpty(seqCellStr) == true)
                        continue;

                    if (int.TryParse(seqCellStr, out seq) == false)
                    {
                        errorString = $"Excel文件{AppConst.OVERVIEW_SHEET_NAME}表中第{rowIndex + 1}行的序号非法，单元格值为{seqCellStr}";
                        return null;
                    }
                    string remark = (row.GetCell(5) != null) ? row.GetCell(5).ToString().Trim() : string.Empty;
                    foreach (string illegalStr in AppConst.FILE_NAME_ILLEGAL_STR)
                    {
                        if (remark.Contains(illegalStr))
                        {
                            errorString = $"Excel文件{AppConst.OVERVIEW_SHEET_NAME}表中第{rowIndex + 1}行要作为文件夹名的备注内容非法，因为在Windows系统中的文件名不能含有{illegalStr}，请修改Excel文件后重试";
                            return null;
                        }
                    }
                    // 对应原本存档文件夹的序号（逆序时不同于Excel中的编号）
                    int matchToSaverFolderSeq = (isRevertSaverNumberSeq == true ? maxRecordSeq - seq + 1 : seq);
                    if (info.SeqToNewFolderName.ContainsKey(matchToSaverFolderSeq))
                    {
                        errorString = $"Excel文件{AppConst.OVERVIEW_SHEET_NAME}表出现了相同的序号{seq}，请不要手工修改除了备注列以外的任何内容，否则本工具无法正常运行";
                        return null;
                    }
                    info.SeqToNewFolderName.Add(matchToSaverFolderSeq, string.IsNullOrEmpty(remark) ? null : remark);
                }
            }

            info.MaxRecordSeq = maxRecordSeq;
            errorString = null;
            return info;
        }
    }
}
