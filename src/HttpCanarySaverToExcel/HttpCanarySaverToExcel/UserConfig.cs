using System.Collections.Generic;
using System.IO;

namespace HttpCanarySaverToExcel
{
    public class UserConfig
    {
        // 选择的HttpCanary存档文件夹（里面为各个顺序数字编号的请求文件夹）
        public string SaverFolderPath { get; set; }
        // 选择的要导出保存的Excel文件存放路径
        public string ExportExcelPath { get; set; }
        // 是否将存档编号进行逆序，以便和HttpCanary中显示的顺序保持一致
        public bool IsRevertSaverNumberSeq { get; set; }
        // 如果上面选择了进行逆序，可以指定最大的纪录编号，为0表示不指定
        public int InputMaxRecordSeq { get; set; }
        // 只分析以下App（包名）的通讯（null表示不进行筛选）
        public List<string> TargetAppPackageNames { get; set; }
        // 忽略以下请求Header中的键名（以全小写存储）
        public List<string> IgnoreReqHeaderName { get; set; } = new List<string>();
        // 忽略以下响应Header中的键名（以全小写存储）
        public List<string> IgnoreRespHeaderName { get; set; } = new List<string>();
        // 是否对URL进行解码显示
        public bool IsUrlDecode { get; set; }
        // 是否对Json进行格式化输出
        public bool IsPrettyPrintJson { get; set; }

        /// <summary>
        /// 检查用户配置是否正确
        /// </summary>
        /// <param name="errorString">如果检查出错误，返回错误信息</param>
        /// <returns>用户配置是否检查无误</returns>
        public bool CheckConfig(out string errorString)
        {
            if (string.IsNullOrEmpty(SaverFolderPath) == true)
            {
                errorString = "未选择HttpCanary存档文件夹";
                return false;
            }
            if (Directory.Exists(SaverFolderPath) == false)
            {
                errorString = "选择的HttpCanary存档文件夹不存在";
                return false;
            }
            if (string.IsNullOrEmpty(ExportExcelPath) == true)
            {
                errorString = "未选择要导出保存的Excel文件存放路径";
                return false;
            }
            string exportExcelFolderPath = Path.GetDirectoryName(ExportExcelPath);
            if (Directory.Exists(exportExcelFolderPath) == false)
            {
                errorString = "选择要导出保存的Excel文件所在文件夹不存在";
                return false;
            }
            string fileExtension = Path.GetExtension(ExportExcelPath);
            if (".xlsx".Equals(fileExtension, System.StringComparison.CurrentCultureIgnoreCase) == false)
            {
                errorString = "要导出保存的Excel文件扩展名必须为xlsx";
                return false;
            }
            if (IsRevertSaverNumberSeq == true)
            {
                if (InputMaxRecordSeq < 0)
                {
                    errorString = "指定的最大记录编号非法，请输入大于0的整数编号。若最大编号的文件夹在存档文件夹之中，可以自动识别，不必填写";
                    return false;
                }
            }

            errorString = null;
            return true;
        }
    }
}
