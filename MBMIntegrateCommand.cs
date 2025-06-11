// ✅ MBMIntegrateCommand.cs (최종 실행 엔트리)
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ConnectorSizeExport.Helpers;
using ConnectorSizeExport.IO;
using ConnectorSizeExport.Models;
using ConnectorSizeExport.Modules;
using ConnectorSizeExport.Settings;
using System;
using System.IO;
using System.Linq;

namespace ConnectorSizeExport.Commands
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class MBMIntegrateCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var connectorRows = ConnectorExportReader.ReadExcel(IntegrateSettings.ConnectorExportPath);
                var settingConditions = SettingConditionReader.Read(IntegrateSettings.SettingPath);

                // ✅ A~I 조건 필터 결과 CSV로 저장
                var debugAtoIPath = @"C:\Temp\DebugAtoI.xlsx";
                DebugAtoIConditionExporter.Export(debugAtoIPath, connectorRows, settingConditions);


                // ✅ LargeDiameter 범위 조건 검증용 CSV 저장
                var debugLargePath = @"C:\Temp\DebugLargeDiameterCheck.xlsx";
                DebugLargeDiameterCheck.Export(debugLargePath, connectorRows, settingConditions);



                var filtered = connectorRows
                    .Where(row => settingConditions.Any(cond =>
                        (string.IsNullOrWhiteSpace(cond.BMArea) || SettingComparer.IsFieldMatch(cond.BMArea, row.BMArea)) &&
                        (string.IsNullOrWhiteSpace(cond.BMUnit) || SettingComparer.IsFieldMatch(cond.BMUnit, row.BMUnit)) &&
                        (string.IsNullOrWhiteSpace(cond.BMZone) || SettingComparer.IsFieldMatch(cond.BMZone, row.BMZone)) &&
                        (string.IsNullOrWhiteSpace(cond.BMDiscipline) || SettingComparer.IsFieldMatch(cond.BMDiscipline, row.BMDiscipline)) &&
                        (string.IsNullOrWhiteSpace(cond.BMSubDiscipline) || SettingComparer.IsFieldMatch(cond.BMSubDiscipline, row.BMSubDiscipline)) &&
                        (string.IsNullOrWhiteSpace(cond.SystemType) || SettingComparer.IsFieldMatch(cond.SystemType, row.SystemType)) &&
                        (string.IsNullOrWhiteSpace(cond.BMFluid) || SettingComparer.IsFieldMatch(cond.BMFluid, row.BMFluid)) &&
                        (string.IsNullOrWhiteSpace(cond.BMClass) || SettingComparer.IsFieldMatch(cond.BMClass, row.BMClass)) &&
                        (string.IsNullOrWhiteSpace(cond.BMScode) || SettingComparer.IsFieldMatch(cond.BMScode, row.BMScode))
                    ))
                    .Select(row => IntegrateRowBuilder.Build(row, settingConditions))
                    .ToList();

                if (filtered.Count == 0)
                {
                    TaskDialog.Show("MBM Integrate", "조건을 만족하는 항목이 없습니다.");
                    return Result.Succeeded;
                }

                var integrateFilePath = IntegrateSettings.GetIntegrateExportPath();
                SheetWriter.WriteIntegrateSheet(integrateFilePath, filtered);
                DebugLargeDiameterCheck.Export(integrateFilePath, connectorRows, settingConditions);

                TaskDialog.Show("MBM Integrate", $"완료되었습니다.\n조건 충족 항목 수: {filtered.Count}");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
