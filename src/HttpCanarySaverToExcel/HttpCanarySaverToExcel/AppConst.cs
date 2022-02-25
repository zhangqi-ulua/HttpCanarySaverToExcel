using System.Globalization;

namespace HttpCanarySaverToExcel
{
    public class AppConst
    {
        // Windows系统中文件名中不能出现的字符
        public static readonly string[] FILE_NAME_ILLEGAL_STR = new string[] { "\\", "/", ":", "*", "?", "\"", "<", ">", "|" };
        /**
         * 解析HttpCanary存档文件中的时间类常量
         */
        public static readonly CultureInfo EN_CULTURE_INFO = CultureInfo.CreateSpecificCulture("en-US");
        public static readonly string GMT_TIME_FORMAT = "ddd, dd-MMM-yyyy HH:mm:ss zz";
        /**
         * 生成Excel文件时的常量
         */
        // Excel一个单元格中文本最大长度为32767，本工具对超过以下字符数的单元格（目前仅可能是打印服务器响应内容的单元格）不显示全部内容
        public static readonly int MAX_SHOW_TEXT_LENGTH_IN_CELL = 32500;
        // 对上面设置的超出最大字符数的单元格，保留上面的字符数后，在开头追加的提示文本
        public static readonly string OVERFLOW_TEXT_TIPS = "【超出单元格文本限制，无法完全显示】";
        // 总览Sheet表的名称
        public static readonly string OVERVIEW_SHEET_NAME = "总览";
        // 总览Sheet表中各字段的标题
        public static readonly string[] OVERVIEW_COLUMN_TITLES = { "序号", "App包名", "协议类型", "URL", "跳转到详情Sheet", "备注" };
        // 总览Sheet表中各字段所在列的列宽
        public static readonly float[] OVERVIEW_SHEET_COLUMN_WIDTH = { 5.5f, 45f, 8f, 70f, 24.5f, 77f };
        // 各个详情Sheet表中各字段所在列的列宽
        public static readonly float[] DETIAL_SHEET_COLUMN_WIDTH = { 32.5f, 200f };
        // 从哪行index开始为各条记录的OneOverviewRecord
        public static readonly int START_RECORD_ROW_INDEX = 3;
        // 逆序与非逆序时总览中注意事项的文本
        public static readonly string IS_REVERT_SAVER_NUMBER_SEQ_NOTICE_TEXT = "序号与App中显示一致，而与存档文件夹中编号相反";
        public static readonly string NOT_REVERT_SAVER_NUMBER_SEQ_NOTICE_TEXT = "序号与存档文件夹中一致，而与App中显示相反";
    }
}
