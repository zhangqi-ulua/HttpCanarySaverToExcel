using Microsoft.VisualBasic.Devices;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace HttpCanarySaverToExcel
{
    public partial class MainForm : Form
    {
        private bool IsShowExtraTools = false;
        // key:要重命名的某个文件夹的完整路径, value:要改为的新文件名（不含路径）
        Dictionary<string, string> TodoRenameInfo = null;

        public MainForm()
        {
            InitializeComponent();

            ChangeExtraToolsShow(IsShowExtraTools);
        }

        private void ChangeExtraToolsShow(bool isShow)
        {
            if (isShow == true)
            {
                BtnIsShowExtraTools.Text = "<< 收起附加工具";
                this.Width = 1314;
                TxtSaverFolderPath.BackColor = Color.Yellow;
                BtnChooseSaverFolder.BackColor = Color.Yellow;
                ChkIsRevertSaverNumberSeq.BackColor = Color.Yellow;
                TxtInputMaxSeq.BackColor = Color.Yellow;
            }
            else
            {
                BtnIsShowExtraTools.Text = "展开附加工具 >>";
                this.Width = 615;
                TxtSaverFolderPath.BackColor = SystemColors.Window;
                BtnChooseSaverFolder.BackColor = SystemColors.Control;
                ChkIsRevertSaverNumberSeq.BackColor = SystemColors.Control;
                TxtInputMaxSeq.BackColor = SystemColors.Window;
            }
        }

        private void BtnChooseSaverFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择HttpCanary存档文件夹";
            dialog.ShowNewFolderButton = false;
            if (dialog.ShowDialog() == DialogResult.OK)
                TxtSaverFolderPath.Text = dialog.SelectedPath;
        }

        private void BtnExportExcelPath_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Title = "请选择要导出保存的Excel文件存放路径";
            dialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
                TxtExportExcelPath.Text = dialog.FileName;
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            UserConfig userConfig = GenerateUserConfig();
            string errorString = null;
            if (userConfig.CheckConfig(out errorString) == false)
            {
                MessageBox.Show(this, errorString, "请修正错误后重试", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (ExportUtil.ExecuteExport(userConfig, out errorString) == false)
                MessageBox.Show(this, errorString, "导出失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                MessageBox.Show(this, $"成功导出至：\n{userConfig.ExportExcelPath}", "导出成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // 打开资源管理器并定位到生成的Excel文件
                System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo("Explorer.exe");
                psi.Arguments = "/e,/select," + userConfig.ExportExcelPath;
                System.Diagnostics.Process.Start(psi);
            }
        }

        private UserConfig GenerateUserConfig()
        {
            UserConfig userConfig = new UserConfig();
            userConfig.SaverFolderPath = TxtSaverFolderPath.Text.Trim();
            userConfig.ExportExcelPath = TxtExportExcelPath.Text.Trim();
            userConfig.IsRevertSaverNumberSeq = ChkIsRevertSaverNumberSeq.Checked;
            userConfig.IsUrlDecode = ChkIsUrlDecode.Checked;
            userConfig.IsPrettyPrintJson = chkIsPrettyPrintJson.Checked;
            if (userConfig.IsRevertSaverNumberSeq == true)
            {
                string inputMaxRecordSeqString = TxtInputMaxSeq.Text.Trim();
                if (string.IsNullOrEmpty(inputMaxRecordSeqString) == false)
                {
                    int inputMaxRecordSeq = -1;
                    // 注意当TryParse无法转换时，会将out的变量赋值为0，而不是不赋值
                    if (int.TryParse(inputMaxRecordSeqString, out inputMaxRecordSeq) == true)
                        userConfig.InputMaxRecordSeq = inputMaxRecordSeq;
                    else
                        userConfig.InputMaxRecordSeq = -1;
                }
            }
            string inputTxtTargetAppPackageNames = TxtTargetAppPackageNames.Text.Trim();
            if (string.IsNullOrEmpty(inputTxtTargetAppPackageNames) == false)
            {
                string[] tempArray = inputTxtTargetAppPackageNames.Split(new string[] { "|" }, StringSplitOptions.None);
                if (tempArray.Length > 0)
                {
                    userConfig.TargetAppPackageNames = new List<string>();
                    foreach (string str in tempArray)
                    {
                        if (userConfig.TargetAppPackageNames.Contains(str) == false)
                            userConfig.TargetAppPackageNames.Add(str);
                    }
                }
            }
            string inputIgnoreReqHeaderName = TxtIgnoreReqHeaderName.Text.Trim();
            if (string.IsNullOrEmpty(inputIgnoreReqHeaderName) == false)
            {
                // 全转为小写存储，以便之后进行比较
                inputIgnoreReqHeaderName = inputIgnoreReqHeaderName.ToLower();
                string[] tempArray = inputIgnoreReqHeaderName.Split(new string[] { "|" }, StringSplitOptions.None);
                foreach (string str in tempArray)
                {
                    if (userConfig.IgnoreReqHeaderName.Contains(str) == false)
                        userConfig.IgnoreReqHeaderName.Add(str);
                }
            }
            string inputIgnoreResqHeaderName = TxtIgnoreRespHeaderName.Text.Trim();
            if (string.IsNullOrEmpty(inputIgnoreResqHeaderName) == false)
            {
                // 全转为小写存储，以便之后进行比较
                inputIgnoreResqHeaderName = inputIgnoreResqHeaderName.ToLower();
                string[] tempArray = inputIgnoreResqHeaderName.Split(new string[] { "|" }, StringSplitOptions.None);
                foreach (string str in tempArray)
                {
                    if (userConfig.IgnoreRespHeaderName.Contains(str) == false)
                        userConfig.IgnoreRespHeaderName.Add(str);
                }
            }

            return userConfig;
        }

        private void ChkIsRevertSaverNumberSeq_CheckedChanged(object sender, EventArgs e)
        {
            LblInputMaxSeqTips.Enabled = ChkIsRevertSaverNumberSeq.Checked;
            TxtInputMaxSeq.Enabled = ChkIsRevertSaverNumberSeq.Checked;
        }

        private void BtnIsShowExtraTools_Click(object sender, EventArgs e)
        {
            IsShowExtraTools = !IsShowExtraTools;
            ChangeExtraToolsShow(IsShowExtraTools);
        }

        private void BtnChooseRemarkExcelPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "用本工具之前生成的Excel文件所在路径";
            dialog.Multiselect = false;
            dialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
                TxtRemarkExcelPath.Text = dialog.FileName;
        }

        private void BtnAnalyzeRenameByRemark_Click(object sender, EventArgs e)
        {
            string saverFolderPath = TxtSaverFolderPath.Text.Trim();
            if (string.IsNullOrEmpty(saverFolderPath) == true)
            {
                MessageBox.Show(this, "未选择HttpCanary存档文件夹", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (Directory.Exists(saverFolderPath) == false)
            {
                MessageBox.Show(this, "选择的HttpCanary存档文件夹不存在", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string excelFilePath = TxtRemarkExcelPath.Text.Trim();
            if (string.IsNullOrEmpty(excelFilePath) == true)
            {
                MessageBox.Show(this, "未选择Excel文件路径", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (File.Exists(excelFilePath) == false)
            {
                MessageBox.Show(this, "选择的Excel文件路径不存在", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (".xlsx".Equals(Path.GetExtension(excelFilePath)) == false)
            {
                MessageBox.Show(this, "未正确选择Excel文件", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            bool isRevertSaverNumberSeq = ChkIsRevertSaverNumberSeq.Checked;
            int inputMaxRecordSeq = 0;
            if (isRevertSaverNumberSeq == true)
            {
                string inputMaxRecordSeqString = TxtInputMaxSeq.Text.Trim();
                if (string.IsNullOrEmpty(inputMaxRecordSeqString) == false)
                {
                    // 注意当TryParse无法转换时，会将out的变量赋值为0，而不是不赋值
                    if (int.TryParse(inputMaxRecordSeqString, out inputMaxRecordSeq) == false || inputMaxRecordSeq < 1)
                    {
                        MessageBox.Show(this, "输入的最大纪录编号值非法", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            AnalyzeRenameFolderByExcelRemarkInfo info = null;
            string errorString = null;
            RtxAnalyzeResult.Clear();
            try
            {
                info = RenameFolderByExcelRemarkUtil.AnalyzeRenameFolderByExcelRemark(excelFilePath, saverFolderPath, isRevertSaverNumberSeq, inputMaxRecordSeq, out errorString);
                if (errorString != null)
                    MessageBox.Show(this, $"分析出错，发现错误为：\n{errorString}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // 生成分析报告
                else
                {
                    /**
                     * 判断存档文件夹是否与Excel文件完美匹配（每条记录都能和存档子文件夹一一对应），
                     * 如果不能则需要提醒用户自己判断是否选择对了匹配的Excel文件和存档文件夹
                     */
                    // 存档文件夹中有但Excel文件中没有的seq
                    List<int> onlyInSaverFolderSeq = new List<int>();
                    // Excel文件中有但存档文件夹中没有的seq
                    List<int> onlyInExcelSeq = new List<int>();
                    // 判断存档文件夹的所有seq，是否与Excel文件中记录的一一对应。注意区分是否逆序的情况
                    foreach (int seqInSaverFolder in info.SeqToOldFolderName.Keys)
                    {
                        int matchExcelSeq = isRevertSaverNumberSeq == true ? info.MaxRecordSeq - seqInSaverFolder + 1 : seqInSaverFolder;
                        if (info.SeqToNewFolderName.ContainsKey(matchExcelSeq) == false)
                            onlyInSaverFolderSeq.Add(matchExcelSeq);
                    }
                    foreach (int seqInExcel in info.SeqToNewFolderName.Keys)
                    {
                        int matchSaverFolderSeq = isRevertSaverNumberSeq == true ? info.MaxRecordSeq - seqInExcel + 1 : seqInExcel;
                        if (info.SeqToOldFolderName.ContainsKey(matchSaverFolderSeq) == false)
                            onlyInExcelSeq.Add(matchSaverFolderSeq);
                    }

                    if (onlyInSaverFolderSeq.Count == 0 && onlyInExcelSeq.Count == 0)
                    {
                        string matchText = "经过检测，存档文件夹与Excel文件中序号完全匹配，可以基本放心地进行重命名\n\n";
                        RtxAnalyzeResult.AppendText(matchText);
                        RtxAnalyzeResult.Select(0, RtxAnalyzeResult.Text.Length);
                        RtxAnalyzeResult.SelectionColor = Color.Green;
                    }
                    else
                    {
                        StringBuilder notMatchTextBuilder = new StringBuilder();
                        notMatchTextBuilder.Append("经过检测，存档文件夹与Excel文件中序号未能完全匹配，请再次确认该存档文件夹是否真的对应该Excel文件，具体不匹配情况如下：\n");
                        if (onlyInSaverFolderSeq.Count > 0)
                        {
                            notMatchTextBuilder.Append($"以下编号只存在于存档文件夹中，而在Excel文件{AppConst.OVERVIEW_SHEET_NAME}表中不存在：\n");
                            notMatchTextBuilder.Append(IntListToStringByASC(onlyInSaverFolderSeq)).Append("\n");
                        }
                        if (onlyInExcelSeq.Count > 0)
                        {
                            notMatchTextBuilder.Append($"以下编号只存Excel文件{AppConst.OVERVIEW_SHEET_NAME}表中，而在存档文件夹中不存在：\n");
                            notMatchTextBuilder.Append(IntListToStringByASC(onlyInExcelSeq)).Append("\n");
                        }
                        notMatchTextBuilder.Append("\n");
                        RtxAnalyzeResult.AppendText(notMatchTextBuilder.ToString());
                        RtxAnalyzeResult.Select(0, RtxAnalyzeResult.Text.Length);
                        RtxAnalyzeResult.SelectionColor = Color.Red;
                    }

                    /**
                     * 分析出需要进行重命名的存档文件夹
                     */
                    TodoRenameInfo = new Dictionary<string, string>();
                    StringBuilder needRenameInfoBuilder = new StringBuilder();

                    foreach (int seq in info.SeqToOldFolderName.Keys)
                    {
                        if (info.SeqToNewFolderName.ContainsKey(seq))
                        {
                            // 提取出原来在文件夹名上设置的备注
                            string oldRemark = null;
                            string folderPath = info.SeqToOldFolderName[seq];
                            int lastIndexForwardSlash = folderPath.LastIndexOf("\\");
                            string folderName = lastIndexForwardSlash > -1 && lastIndexForwardSlash != folderPath.Length - 1 ? folderPath.Substring(lastIndexForwardSlash + 1) : folderPath;
                            int firstIndexForBracket = folderName.IndexOf("（");
                            if (firstIndexForBracket != -1)
                            {
                                int lastIndexForBracket = folderName.LastIndexOf("）");
                                oldRemark = folderName.Substring(firstIndexForBracket + 1, lastIndexForBracket - firstIndexForBracket - 1);
                            }
                            string newRemark = info.SeqToNewFolderName[seq];

                            if ((oldRemark == null && newRemark != null) || (oldRemark != null && oldRemark.Equals(newRemark) == false))
                            {
                                // 如果没有备注，也就不用添加括号
                                string newName = (string.IsNullOrEmpty(newRemark) ? seq.ToString() : string.Concat(seq, "（", newRemark, "）"));
                                TodoRenameInfo.Add(folderPath, newName);
                                needRenameInfoBuilder.Append($"{folderName}\n{newName}\n\n");
                            }
                        }
                    }

                    needRenameInfoBuilder.Insert(0, $"共有{info.SeqToOldFolderName.Count}个存档文件夹，其中{TodoRenameInfo.Count}个本次需要进行重命名\n{(TodoRenameInfo.Count > 0 ? "下面每组两行分别为原文件名和本次要修改为的新文件名：\n\n" : string.Empty)}");

                    RtxAnalyzeResult.AppendText(needRenameInfoBuilder.ToString());
                    BtnRenameByRemark.Enabled = true;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, $"分析出错，抛出异常为：\n{exception.ToString()}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        /// <summary>
        /// 将一个int列表的元素进行升序输出
        /// </summary>
        /// <param name="list">要进行输出的列表</param>
        /// <param name="splitStr">数字之间的分隔符，不传入默认为一个英文逗号</param>
        /// <returns></returns>
        private string IntListToStringByASC(List<int> list, string splitStr = ",")
        {
            StringBuilder sb = new StringBuilder();
            list.Sort();
            foreach (int num in list)
                sb.Append(num).Append(splitStr);

            sb.Remove(sb.Length - 1, 1);
            return sb.ToString();
        }

        /// <summary>
        /// 执行重命名，因为某个seq开头的文件夹必然只有一个，所以依次执行重命名不会出现文件名重复冲突的情况
        /// </summary>
        private void BtnRenameByRemark_Click(object sender, EventArgs e)
        {
            if (TodoRenameInfo.Count < 1)
            {
                MessageBox.Show(this, "没有任何需要重命名的存档文件夹", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            AppendTextToRtx("\n开始执行重命名：\n");
            try
            {
                Computer myComputer = new Computer();

                foreach (var keyAndValuePair in TodoRenameInfo)
                {
                    string folderFullPath = keyAndValuePair.Key;
                    string newName = keyAndValuePair.Value;

                    myComputer.FileSystem.RenameDirectory(folderFullPath, newName);
                    int lastIndexForwardSlash = folderFullPath.LastIndexOf("\\");
                    string folderName = lastIndexForwardSlash > -1 && lastIndexForwardSlash != folderFullPath.Length - 1 ? folderFullPath.Substring(lastIndexForwardSlash + 1) : folderFullPath;
                    AppendTextToRtx($"完成：{folderName} => {newName}\n");
                }

                AppendTextToRtx("\n执行成功\n");
                MessageBox.Show(this, "重命名执行完毕", "恭喜", MessageBoxButtons.OK, MessageBoxIcon.Information);
                BtnRenameByRemark.Enabled = false;
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, $"重命名失败，抛出异常为\n{exception.ToString()}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void AppendTextToRtx(string text)
        {
            // 让富文本框获取焦点
            RtxAnalyzeResult.Focus();
            // 设置光标的位置到文本结尾
            RtxAnalyzeResult.Select(RtxAnalyzeResult.TextLength, 0);
            // 滚动到富文本框光标处
            RtxAnalyzeResult.ScrollToCaret();
            // 追加内容
            RtxAnalyzeResult.AppendText(text);
        }
    }
}
