// ✅ IntegrateSettings.cs - 통합 설정 및 경로 관리
namespace ConnectorSizeExport.Settings
{
    public static class IntegrateSettings
    {
        public static string BasePath => @"C:\\Temp";

        public static string ConnectorExportPath => System.IO.Path.Combine(BasePath, "ConnectorSizes_20250602_141506.xlsx");

        public static string SettingPath => System.IO.Path.Combine(BasePath, "Setting Excel File.xlsx");

        public static string GetIntegrateExportPath() =>
            System.IO.Path.Combine(BasePath, $"ConnectorSizes_Export_Integrate_{System.DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
    }
}