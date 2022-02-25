namespace HttpCanarySaverToExcel
{
    /// <summary>
    /// 对应生成的Excel文件中，总览Sheet中的一条记录
    /// </summary>
    public class OneOverviewRecord
    {
        // 记录对应的文件夹信息
        public OneRecordFileInfo RecordFileInfo { get; set; }
        // 通讯协议
        public NetTypeEnum NetType { get; set; }
        // 请求App对应的包名
        public string AppPackageName { get; set; }
        // 请求的地址
        public string Url { get; set; }
        // 是否有对应的Sheet表分析这条记录详情
        public bool IsGenerateDetialSheet { get; set; }
        // 如果没有对应的Sheet表分析这条记录详情，原因是什么
        public string NotGenerateDetialSheetReason { get; set; }
    }

    public enum NetTypeEnum
    {
        UNIMPLEMENTED,

        HTTP,
        HTTPS,
        TCP,
        UDP,
    }
}
