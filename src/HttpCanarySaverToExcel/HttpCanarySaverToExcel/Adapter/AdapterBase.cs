using NPOI.SS.UserModel;

namespace HttpCanarySaverToExcel.Adapter
{
    public abstract class AdapterBase
    {
        // 哪个App发起的请求（用包名表示）
        public string AppPackageName { get; set; }
        // 请求服务器的IP和端口（格式为“IP:端口号”）
        public string RemoteIpAndPort { get; set; }

        /// <summary>
        /// 将本Adapter对应的详情信息写为Excel的对应Sheet表
        /// </summary>
        /// <param name="workbook">Excel工作簿文件</param>
        /// <param name="overviewRecord">对应的总览信息，含序号作为Sheet表名</param>
        public abstract void WriteExcelSheet(IWorkbook workbook, UserConfig userConfig, OneOverviewRecord overviewRecord);
    }
}
