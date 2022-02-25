
namespace HttpCanarySaverToExcel
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.TxtSaverFolderPath = new System.Windows.Forms.TextBox();
            this.BtnChooseSaverFolder = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.TxtExportExcelPath = new System.Windows.Forms.TextBox();
            this.BtnExportExcelPath = new System.Windows.Forms.Button();
            this.BtnExport = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.ChkIsRevertSaverNumberSeq = new System.Windows.Forms.CheckBox();
            this.LblInputMaxSeqTips = new System.Windows.Forms.Label();
            this.TxtInputMaxSeq = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.TxtIgnoreReqHeaderName = new System.Windows.Forms.TextBox();
            this.TxtIgnoreRespHeaderName = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.TxtTargetAppPackageNames = new System.Windows.Forms.TextBox();
            this.ChkIsUrlDecode = new System.Windows.Forms.CheckBox();
            this.chkIsPrettyPrintJson = new System.Windows.Forms.CheckBox();
            this.GrpRenameFolderByExcelRemark = new System.Windows.Forms.GroupBox();
            this.BtnRenameByRemark = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.RtxAnalyzeResult = new System.Windows.Forms.RichTextBox();
            this.BtnAnalyzeRenameByRemark = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.BtnChooseRemarkExcelPath = new System.Windows.Forms.Button();
            this.TxtRemarkExcelPath = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.BtnIsShowExtraTools = new System.Windows.Forms.Button();
            this.GrpRenameFolderByExcelRemark.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(12, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(569, 78);
            this.label1.TabIndex = 0;
            this.label1.Text = resources.GetString("label1.Text");
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 124);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(401, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "请选择HttpCanary存档文件夹（里面为各个顺序数字编号的请求文件夹）：";
            // 
            // TxtSaverFolderPath
            // 
            this.TxtSaverFolderPath.Location = new System.Drawing.Point(14, 150);
            this.TxtSaverFolderPath.Name = "TxtSaverFolderPath";
            this.TxtSaverFolderPath.Size = new System.Drawing.Size(471, 21);
            this.TxtSaverFolderPath.TabIndex = 2;
            // 
            // BtnChooseSaverFolder
            // 
            this.BtnChooseSaverFolder.Location = new System.Drawing.Point(503, 148);
            this.BtnChooseSaverFolder.Name = "BtnChooseSaverFolder";
            this.BtnChooseSaverFolder.Size = new System.Drawing.Size(75, 23);
            this.BtnChooseSaverFolder.TabIndex = 3;
            this.BtnChooseSaverFolder.Text = "选择";
            this.BtnChooseSaverFolder.UseVisualStyleBackColor = true;
            this.BtnChooseSaverFolder.Click += new System.EventHandler(this.BtnChooseSaverFolder_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 186);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(227, 12);
            this.label3.TabIndex = 4;
            this.label3.Text = "请选择要导出保存的Excel文件存放路径：";
            // 
            // TxtExportExcelPath
            // 
            this.TxtExportExcelPath.Location = new System.Drawing.Point(14, 214);
            this.TxtExportExcelPath.Name = "TxtExportExcelPath";
            this.TxtExportExcelPath.Size = new System.Drawing.Size(471, 21);
            this.TxtExportExcelPath.TabIndex = 5;
            // 
            // BtnExportExcelPath
            // 
            this.BtnExportExcelPath.Location = new System.Drawing.Point(503, 212);
            this.BtnExportExcelPath.Name = "BtnExportExcelPath";
            this.BtnExportExcelPath.Size = new System.Drawing.Size(75, 23);
            this.BtnExportExcelPath.TabIndex = 6;
            this.BtnExportExcelPath.Text = "选择";
            this.BtnExportExcelPath.UseVisualStyleBackColor = true;
            this.BtnExportExcelPath.Click += new System.EventHandler(this.BtnExportExcelPath_Click);
            // 
            // BtnExport
            // 
            this.BtnExport.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.BtnExport.Location = new System.Drawing.Point(206, 697);
            this.BtnExport.Name = "BtnExport";
            this.BtnExport.Size = new System.Drawing.Size(182, 63);
            this.BtnExport.TabIndex = 7;
            this.BtnExport.Text = "导  出";
            this.BtnExport.UseVisualStyleBackColor = true;
            this.BtnExport.Click += new System.EventHandler(this.BtnExport_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 258);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 8;
            this.label4.Text = "高级选项：";
            // 
            // ChkIsRevertSaverNumberSeq
            // 
            this.ChkIsRevertSaverNumberSeq.Checked = true;
            this.ChkIsRevertSaverNumberSeq.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ChkIsRevertSaverNumberSeq.Location = new System.Drawing.Point(14, 283);
            this.ChkIsRevertSaverNumberSeq.Name = "ChkIsRevertSaverNumberSeq";
            this.ChkIsRevertSaverNumberSeq.Size = new System.Drawing.Size(567, 35);
            this.ChkIsRevertSaverNumberSeq.TabIndex = 9;
            this.ChkIsRevertSaverNumberSeq.Text = "逆序存档请求的序号（HttpCanary导出的存档文件夹序号与App中所见序号正好相反，勾选后可使Excel中请求编号与App一致）";
            this.ChkIsRevertSaverNumberSeq.UseVisualStyleBackColor = true;
            this.ChkIsRevertSaverNumberSeq.CheckedChanged += new System.EventHandler(this.ChkIsRevertSaverNumberSeq_CheckedChanged);
            // 
            // LblInputMaxSeqTips
            // 
            this.LblInputMaxSeqTips.Location = new System.Drawing.Point(57, 331);
            this.LblInputMaxSeqTips.Name = "LblInputMaxSeqTips";
            this.LblInputMaxSeqTips.Size = new System.Drawing.Size(345, 29);
            this.LblInputMaxSeqTips.TabIndex = 10;
            this.LblInputMaxSeqTips.Text = "逆序选项下，可指定最大的记录编号，以便与App实际对应，若最大编号的文件夹在存档文件夹之中，可以自动识别，此处留空";
            // 
            // TxtInputMaxSeq
            // 
            this.TxtInputMaxSeq.Location = new System.Drawing.Point(408, 335);
            this.TxtInputMaxSeq.Name = "TxtInputMaxSeq";
            this.TxtInputMaxSeq.Size = new System.Drawing.Size(100, 21);
            this.TxtInputMaxSeq.TabIndex = 11;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 446);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(419, 12);
            this.label5.TabIndex = 12;
            this.label5.Text = "忽略掉Http请求Header中的以下字段（用|分隔各个字段名，不分区大小写）：";
            // 
            // TxtIgnoreReqHeaderName
            // 
            this.TxtIgnoreReqHeaderName.Location = new System.Drawing.Point(14, 474);
            this.TxtIgnoreReqHeaderName.Multiline = true;
            this.TxtIgnoreReqHeaderName.Name = "TxtIgnoreReqHeaderName";
            this.TxtIgnoreReqHeaderName.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.TxtIgnoreReqHeaderName.Size = new System.Drawing.Size(564, 45);
            this.TxtIgnoreReqHeaderName.TabIndex = 13;
            this.TxtIgnoreReqHeaderName.Text = "Accept-Encoding|Connection|Content-Length";
            // 
            // TxtIgnoreRespHeaderName
            // 
            this.TxtIgnoreRespHeaderName.Location = new System.Drawing.Point(14, 566);
            this.TxtIgnoreRespHeaderName.Multiline = true;
            this.TxtIgnoreRespHeaderName.Name = "TxtIgnoreRespHeaderName";
            this.TxtIgnoreRespHeaderName.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.TxtIgnoreRespHeaderName.Size = new System.Drawing.Size(564, 45);
            this.TxtIgnoreRespHeaderName.TabIndex = 15;
            this.TxtIgnoreRespHeaderName.Text = "Connection";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(12, 538);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(419, 12);
            this.label6.TabIndex = 14;
            this.label6.Text = "忽略掉Http响应Header中的以下字段（用|分隔各个字段名，不分区大小写）：";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(12, 382);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(407, 12);
            this.label7.TabIndex = 16;
            this.label7.Text = "只分析以下App（包名）的通讯（用|分隔各个包名，留空则分析全部App）：";
            // 
            // TxtTargetAppPackageNames
            // 
            this.TxtTargetAppPackageNames.Location = new System.Drawing.Point(12, 404);
            this.TxtTargetAppPackageNames.Multiline = true;
            this.TxtTargetAppPackageNames.Name = "TxtTargetAppPackageNames";
            this.TxtTargetAppPackageNames.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.TxtTargetAppPackageNames.Size = new System.Drawing.Size(564, 30);
            this.TxtTargetAppPackageNames.TabIndex = 17;
            // 
            // ChkIsUrlDecode
            // 
            this.ChkIsUrlDecode.AutoSize = true;
            this.ChkIsUrlDecode.Checked = true;
            this.ChkIsUrlDecode.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ChkIsUrlDecode.Location = new System.Drawing.Point(14, 630);
            this.ChkIsUrlDecode.Name = "ChkIsUrlDecode";
            this.ChkIsUrlDecode.Size = new System.Drawing.Size(462, 16);
            this.ChkIsUrlDecode.TabIndex = 18;
            this.ChkIsUrlDecode.Text = "对HTTP请求的URL进行解码显示（例如：“%E5%BC%A0%E9%BD%90”解码为“张齐”）";
            this.ChkIsUrlDecode.UseVisualStyleBackColor = true;
            // 
            // chkIsPrettyPrintJson
            // 
            this.chkIsPrettyPrintJson.AutoSize = true;
            this.chkIsPrettyPrintJson.Checked = true;
            this.chkIsPrettyPrintJson.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIsPrettyPrintJson.Location = new System.Drawing.Point(14, 658);
            this.chkIsPrettyPrintJson.Name = "chkIsPrettyPrintJson";
            this.chkIsPrettyPrintJson.Size = new System.Drawing.Size(180, 16);
            this.chkIsPrettyPrintJson.TabIndex = 19;
            this.chkIsPrettyPrintJson.Text = "对Json字符串进行格式化输出";
            this.chkIsPrettyPrintJson.UseVisualStyleBackColor = true;
            // 
            // GrpRenameFolderByExcelRemark
            // 
            this.GrpRenameFolderByExcelRemark.Controls.Add(this.BtnRenameByRemark);
            this.GrpRenameFolderByExcelRemark.Controls.Add(this.label10);
            this.GrpRenameFolderByExcelRemark.Controls.Add(this.RtxAnalyzeResult);
            this.GrpRenameFolderByExcelRemark.Controls.Add(this.BtnAnalyzeRenameByRemark);
            this.GrpRenameFolderByExcelRemark.Controls.Add(this.label9);
            this.GrpRenameFolderByExcelRemark.Controls.Add(this.BtnChooseRemarkExcelPath);
            this.GrpRenameFolderByExcelRemark.Controls.Add(this.TxtRemarkExcelPath);
            this.GrpRenameFolderByExcelRemark.Controls.Add(this.label8);
            this.GrpRenameFolderByExcelRemark.Location = new System.Drawing.Point(619, 20);
            this.GrpRenameFolderByExcelRemark.Name = "GrpRenameFolderByExcelRemark";
            this.GrpRenameFolderByExcelRemark.Size = new System.Drawing.Size(663, 740);
            this.GrpRenameFolderByExcelRemark.TabIndex = 20;
            this.GrpRenameFolderByExcelRemark.TabStop = false;
            this.GrpRenameFolderByExcelRemark.Text = "将生成的Excel文件中填写的每个抓包的备注同步修改到存档文件夹";
            // 
            // BtnRenameByRemark
            // 
            this.BtnRenameByRemark.Enabled = false;
            this.BtnRenameByRemark.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.BtnRenameByRemark.Location = new System.Drawing.Point(265, 678);
            this.BtnRenameByRemark.Name = "BtnRenameByRemark";
            this.BtnRenameByRemark.Size = new System.Drawing.Size(132, 45);
            this.BtnRenameByRemark.TabIndex = 11;
            this.BtnRenameByRemark.Text = "执行重命名";
            this.BtnRenameByRemark.UseVisualStyleBackColor = true;
            this.BtnRenameByRemark.Click += new System.EventHandler(this.BtnRenameByRemark_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(20, 225);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(197, 12);
            this.label10.TabIndex = 10;
            this.label10.Text = "以下显示分析出的要重命名的结果：";
            // 
            // RtxAnalyzeResult
            // 
            this.RtxAnalyzeResult.Location = new System.Drawing.Point(22, 254);
            this.RtxAnalyzeResult.Name = "RtxAnalyzeResult";
            this.RtxAnalyzeResult.Size = new System.Drawing.Size(624, 405);
            this.RtxAnalyzeResult.TabIndex = 9;
            this.RtxAnalyzeResult.Text = "";
            // 
            // BtnAnalyzeRenameByRemark
            // 
            this.BtnAnalyzeRenameByRemark.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.BtnAnalyzeRenameByRemark.Location = new System.Drawing.Point(265, 146);
            this.BtnAnalyzeRenameByRemark.Name = "BtnAnalyzeRenameByRemark";
            this.BtnAnalyzeRenameByRemark.Size = new System.Drawing.Size(132, 45);
            this.BtnAnalyzeRenameByRemark.TabIndex = 8;
            this.BtnAnalyzeRenameByRemark.Text = "分析重命名";
            this.BtnAnalyzeRenameByRemark.UseVisualStyleBackColor = true;
            this.BtnAnalyzeRenameByRemark.Click += new System.EventHandler(this.BtnAnalyzeRenameByRemark_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.BackColor = System.Drawing.Color.Yellow;
            this.label9.Location = new System.Drawing.Point(20, 34);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(341, 24);
            this.label9.TabIndex = 5;
            this.label9.Text = "请先在左边正确选择或填写以黄色背景突出显示的项目，\r\n特别注意是否逆序的选择与生成时必须一致，否则无法正确对应";
            // 
            // BtnChooseRemarkExcelPath
            // 
            this.BtnChooseRemarkExcelPath.Location = new System.Drawing.Point(571, 103);
            this.BtnChooseRemarkExcelPath.Name = "BtnChooseRemarkExcelPath";
            this.BtnChooseRemarkExcelPath.Size = new System.Drawing.Size(75, 23);
            this.BtnChooseRemarkExcelPath.TabIndex = 4;
            this.BtnChooseRemarkExcelPath.Text = "选择";
            this.BtnChooseRemarkExcelPath.UseVisualStyleBackColor = true;
            this.BtnChooseRemarkExcelPath.Click += new System.EventHandler(this.BtnChooseRemarkExcelPath_Click);
            // 
            // TxtRemarkExcelPath
            // 
            this.TxtRemarkExcelPath.Location = new System.Drawing.Point(22, 105);
            this.TxtRemarkExcelPath.Name = "TxtRemarkExcelPath";
            this.TxtRemarkExcelPath.Size = new System.Drawing.Size(529, 21);
            this.TxtRemarkExcelPath.TabIndex = 3;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(20, 80);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(437, 12);
            this.label8.TabIndex = 0;
            this.label8.Text = "请再选择修改好备注信息的Excel表格（一定不要修改本工具生成的Excel格式）：";
            // 
            // BtnIsShowExtraTools
            // 
            this.BtnIsShowExtraTools.Location = new System.Drawing.Point(456, 711);
            this.BtnIsShowExtraTools.Name = "BtnIsShowExtraTools";
            this.BtnIsShowExtraTools.Size = new System.Drawing.Size(122, 37);
            this.BtnIsShowExtraTools.TabIndex = 21;
            this.BtnIsShowExtraTools.Text = "展开附加工具 >>";
            this.BtnIsShowExtraTools.UseVisualStyleBackColor = true;
            this.BtnIsShowExtraTools.Click += new System.EventHandler(this.BtnIsShowExtraTools_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1294, 776);
            this.Controls.Add(this.BtnIsShowExtraTools);
            this.Controls.Add(this.GrpRenameFolderByExcelRemark);
            this.Controls.Add(this.chkIsPrettyPrintJson);
            this.Controls.Add(this.ChkIsUrlDecode);
            this.Controls.Add(this.TxtTargetAppPackageNames);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.TxtIgnoreRespHeaderName);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.TxtIgnoreReqHeaderName);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.TxtInputMaxSeq);
            this.Controls.Add(this.LblInputMaxSeqTips);
            this.Controls.Add(this.ChkIsRevertSaverNumberSeq);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.BtnExport);
            this.Controls.Add(this.BtnExportExcelPath);
            this.Controls.Add(this.TxtExportExcelPath);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.BtnChooseSaverFolder);
            this.Controls.Add(this.TxtSaverFolderPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "HttpCanary存档转Excel文件 by 张齐 （https://github.com/zhangqi-ulua）";
            this.GrpRenameFolderByExcelRemark.ResumeLayout(false);
            this.GrpRenameFolderByExcelRemark.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TxtSaverFolderPath;
        private System.Windows.Forms.Button BtnChooseSaverFolder;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox TxtExportExcelPath;
        private System.Windows.Forms.Button BtnExportExcelPath;
        private System.Windows.Forms.Button BtnExport;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox ChkIsRevertSaverNumberSeq;
        private System.Windows.Forms.Label LblInputMaxSeqTips;
        private System.Windows.Forms.TextBox TxtInputMaxSeq;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox TxtIgnoreReqHeaderName;
        private System.Windows.Forms.TextBox TxtIgnoreRespHeaderName;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox TxtTargetAppPackageNames;
        private System.Windows.Forms.CheckBox ChkIsUrlDecode;
        private System.Windows.Forms.CheckBox chkIsPrettyPrintJson;
        private System.Windows.Forms.GroupBox GrpRenameFolderByExcelRemark;
        private System.Windows.Forms.Button BtnIsShowExtraTools;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button BtnAnalyzeRenameByRemark;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button BtnChooseRemarkExcelPath;
        private System.Windows.Forms.TextBox TxtRemarkExcelPath;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.RichTextBox RtxAnalyzeResult;
        private System.Windows.Forms.Button BtnRenameByRemark;
    }
}

