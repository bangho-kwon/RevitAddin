using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System.Windows.Interop;
using MTMDCAuth;

namespace PipeInsulationInput
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class InsulationThicknessInput: IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // ✅ 인증 먼저 확인
                string reason;
                if (!ExcelMacAuth.TryAuthorize(out reason))
                {
                    TaskDialog.Show("인증에 실패하였습니다. 관리자에게 문의하세요", reason);
                    return Result.Failed;
                }

                var window = new InsulationSettingsWindow(doc)
                {
                    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                    //Topmost = true
                };

                IntPtr revitHandle = commandData.Application.MainWindowHandle;
                WindowInteropHelper helper = new WindowInteropHelper(window)
                {
                    Owner = revitHandle
                };

                bool? result = window.ShowDialog();
                if (result != true)
                    return Result.Cancelled;

                string selectedCsvPath = window.CsvFilePath;
                bool useBMArea = window.UseBMArea;
                bool useBMUnit = window.UseBMUnit;
                bool useBMZone = window.UseBMZone;
                bool useFluid = window.UseFluid;
                bool useSystemType = window.UseSystemType;

                string bmAreaParam = window.BMAreaParamName;
                string bmUnitParam = window.BMUnitParamName;
                string bmZoneParam = window.BMZoneParamName;
                string fluidParam = window.FluidParamName;
                //string systemTypeParam = window.SystemTypeParamName;

                var pipeCollector = new FilteredElementCollector(doc).OfClass(typeof(Pipe)).WhereElementIsNotElementType();
                var pipeFittingCollector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType();
                var pipeAccessoryCollector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType();

                var allPipeElements = pipeCollector.Cast<Element>()
                    .Concat(pipeFittingCollector.Cast<Element>())
                    .Concat(pipeAccessoryCollector.Cast<Element>()).ToList();

                if (allPipeElements.Count == 0)
                {
                    TaskDialog.Show("오류", "도면에 파이프, 피팅, 액세서리 요소가 존재하지 않습니다.");
                    return Result.Failed;
                }

                bool isMetric = GetPipeSizeUnit(doc);

                using (Transaction t = new Transaction(doc, "파이프 및 관련요소 절연 두께 설정"))
                {
                    t.Start();

                    int updatedCount = 0;
                    int exceptedCount = 0;
                    int resetCount = 0;

                    foreach (var elem in allPipeElements)
                    {
                        // Insulation Except 체크 시 아무 작업도 하지 않음
                        if (IsInsulationExceptChecked(elem))
                        {
                            exceptedCount++;
                            continue;
                        }

                        // 기존 값이 입력되어 있으면 0으로 초기화
                        if (HasValue(elem, "Insulation Fluid Thickness"))
                        {
                            SetParameterValue(elem, "Insulation Fluid Thickness", 0);
                            resetCount++;
                        }

                        // 선택된 파라미터명에 따라 값 읽기
                        string bmArea = NormalizeText(GetParameterValue(elem, bmAreaParam)).ToLower();
                        string bmUnit = NormalizeText(GetParameterValue(elem, bmUnitParam)).ToLower();
                        string bmZone = NormalizeText(GetParameterValue(elem, bmZoneParam)).ToLower();
                        string fluid = NormalizeText(GetParameterValue(elem, fluidParam)).ToLower();
                        string systemType = GetSystemTypeName(elem)?.ToLower();
                        if (string.IsNullOrWhiteSpace(systemType)) systemType = "default";


                        if (string.IsNullOrWhiteSpace(bmUnit)) bmUnit = "default";
                        if (string.IsNullOrWhiteSpace(bmZone)) bmZone = "default";
                        if (string.IsNullOrWhiteSpace(bmArea)) bmArea = "default";

                        double diameter = 0;
                        if (elem is Pipe)
                            diameter = GetDiameter(elem);
                        else if (elem is FamilyInstance || elem is MEPCurve)
                            diameter = GetMaxConnectorDiameter(elem, doc);
                        else
                            diameter = GetDiameter(elem);

                        string resultStr = FindMatchingData(
                            selectedCsvPath, bmArea, bmUnit, bmZone, fluid, systemType, diameter, isMetric,
                            useBMArea, useBMUnit, useBMZone, useFluid, useSystemType);

                        if (int.TryParse(resultStr, out int thickness) && thickness > 0)
                        {
                            SetParameterValue(elem, "Insulation Fluid Thickness", thickness);
                            updatedCount++;
                        }
                    }

                    t.Commit();

                    TaskDialog.Show("결과",
                        $"{updatedCount}개의 요소에 절연 두께가 입력되었습니다.\n" +
                        $"{exceptedCount}개의 요소는 'Insulation Except'가 체크되어 기존 값을 유지하였습니다.\n" +
                        $"{resetCount}개의 요소는 기존 값을 0으로 초기화하였습니다.");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private string FindMatchingData(string filePath, string bmArea, string bmUnit, string bmZone, string fluid, string systemType, double diameter, bool isMetric,
            bool useBMArea, bool useBMUnit, bool useBMZone, bool useFluid, bool useSystemType)
        {
            try
            {
                var lines = File.ReadAllLines(filePath);
                foreach (var line in lines)
                {
                    var values = line.Split(',');
                    if (values.Length < 10) continue;

                    string csvBMArea = NormalizeText(values[0].Trim()).ToLower();
                    string csvBMUnit = NormalizeText(values[1].Trim()).ToLower();
                    string csvBMZone = NormalizeText(values[2].Trim()).ToLower();
                    string csvFluid = NormalizeText(values[3].Trim()).ToLower();
                    string csvSystemType = NormalizeText(values[4].Trim()).ToLower();

                    if (string.IsNullOrWhiteSpace(csvBMArea)) csvBMArea = "default";
                    if (string.IsNullOrWhiteSpace(csvBMUnit)) csvBMUnit = "default";
                    if (string.IsNullOrWhiteSpace(csvBMZone)) csvBMZone = "default";

                    int fromIdx = isMetric ? 5 : 7;
                    int toIdx = isMetric ? 6 : 8;

                    if (double.TryParse(values[fromIdx], out double fromVal) && double.TryParse(values[toIdx], out double toVal))
                    {
                        bool matched = true;

                        if (useBMArea)
                            matched &= MatchValues(bmArea, csvBMArea);
                        if (useBMUnit)
                            matched &= MatchValues(bmUnit, csvBMUnit);
                        if (useBMZone)
                            matched &= MatchValues(bmZone, csvBMZone);
                        if (useFluid)
                            matched &= MatchValues(fluid, csvFluid);
                        if (useSystemType)
                            matched &= MatchValues(systemType, csvSystemType);

                        matched &= (diameter > fromVal && diameter <= toVal);

                        if (matched)
                            return values[9].Trim();
                    }
                }
            }
            catch { }

            return "0";
        }

        private bool MatchValues(string input, string allowed)
        {
            string[] allowedVals = allowed.Split('/');
            string[] inputVals = input.Split('/');

            return inputVals.Any(inputVal => allowedVals.Any(allowedVal =>
                RemoveWhiteSpace(allowedVal).Equals(RemoveWhiteSpace(inputVal), StringComparison.OrdinalIgnoreCase)));
        }

        private string NormalizeText(string input) => input?.Normalize(System.Text.NormalizationForm.FormC) ?? string.Empty;

        private string RemoveWhiteSpace(string input) => string.IsNullOrEmpty(input) ? "" : string.Concat(input.Where(c => !char.IsWhiteSpace(c)));

        private string GetParameterValue(Element element, string paramName)
        {
            var param = element.LookupParameter(paramName);
            return (param == null) ? "" : param.AsString()?.Trim() ?? "";
        }

        private void SetParameterValue(Element element, string paramName, int value)
        {
            var param = element.LookupParameter(paramName);
            if (param != null)
            {
                if (param.StorageType == StorageType.Integer)
                    param.Set(value);
                else if (param.StorageType == StorageType.Double)
                    param.Set(value / 304.8);
                else if (param.StorageType == StorageType.String)
                    param.Set(value.ToString());
            }
        }

        private bool IsInsulationExceptChecked(Element element, string paramName = "Insulation Except")
        {
            var param = element.LookupParameter(paramName);
            return param != null && param.StorageType == StorageType.Integer && param.AsInteger() == 1;
        }

        private bool HasValue(Element element, string paramName)
        {
            var param = element.LookupParameter(paramName);
            if (param == null) return false;
            if (param.StorageType == StorageType.Integer) return param.AsInteger() != 0;
            if (param.StorageType == StorageType.Double) return Math.Abs(param.AsDouble()) > 1e-8;
            if (param.StorageType == StorageType.String) return !string.IsNullOrWhiteSpace(param.AsString()) && param.AsString() != "0";
            return false;
        }

        private bool GetPipeSizeUnit(Document doc)
        {
            Units projectUnits = doc.GetUnits();
            FormatOptions formatOptions = projectUnits.GetFormatOptions(SpecTypeId.PipeSize);
            return formatOptions.GetUnitTypeId() == UnitTypeId.Millimeters;
        }

        private string GetSystemTypeName(Element elem)
        {
            Parameter sysTypeParam = elem.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
            return sysTypeParam?.AsValueString()?.Trim() ?? "";
        }


        private double GetDiameter(Element elem)
        {
            var diameterParam = elem.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
            if (diameterParam != null)
            {
                string diameterRaw = diameterParam.AsValueString()?.Trim() ?? "0";
                string diameterStr = Regex.Replace(diameterRaw, @"[^0-9.]", "");
                if (double.TryParse(diameterStr, out double diameter))
                    return diameter;
            }
            return 0;
        }

        private double GetMaxConnectorDiameter(Element elem, Document doc)
        {
            double maxDiameter = 0.0;
            List<double> diameters = new List<double>();
            var unitType = doc.GetUnits().GetFormatOptions(SpecTypeId.PipeSize).GetUnitTypeId();

            ConnectorManager connectorManager = null;
            if (elem is FamilyInstance fi && fi.MEPModel != null)
                connectorManager = fi.MEPModel.ConnectorManager;
            else if (elem is MEPCurve mc)
                connectorManager = mc.ConnectorManager;

            if (connectorManager == null) return 0.0;

            foreach (Connector conn in connectorManager.Connectors)
            {
                if (conn.Shape == ConnectorProfileType.Round)
                {
                    double diameter = conn.Radius * 2;
                    double converted = UnitUtils.ConvertFromInternalUnits(diameter, unitType);
                    diameters.Add(converted);
                }
            }
            if (diameters.Count > 0)
                maxDiameter = diameters.Max();

            return maxDiameter;
        }
    }
}