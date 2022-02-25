using NPOI.SS.UserModel;

namespace HttpCanarySaverToExcel.Adapter
{
    public class UnimplementedAdapter : AdapterBase
    {
        public override void WriteExcelSheet(IWorkbook workbook, UserConfig userConfig, OneOverviewRecord overviewRecord)
        {
        }
    }
}
