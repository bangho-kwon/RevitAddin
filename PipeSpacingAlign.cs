using MTMDCCommon;
using MTMDCAuth;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Interop;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows;
using System.Text;



namespace PipeSpacingAlign
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AlignAndSetPipeInsulation : IExternalCommand
    {
        private static readonly string SettingsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PipeSpacingAlign",
            "settings.txt");

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // ✅ 인증 먼저 확인
            string reason;
            if (!ExcelMacAuth.TryAuthorize(out reason))
            {
                TaskDialog.Show("인증에 실패하였습니다. 관리자에게 문의하세요", reason);
                return Result.Failed;
            }

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            UIApplication uiApp = commandData.Application;

            LoadSettings(out double spacingMM, out string csvFilePath);


            var wpfWindow = new SpacingInputWindow(spacingMM, csvFilePath);
            var helper = new WindowInteropHelper(wpfWindow);
            helper.Owner = uiApp.MainWindowHandle;

            if (wpfWindow.ShowDialog() != true || !wpfWindow.SpacingMM.HasValue || string.IsNullOrEmpty(wpfWindow.CsvFilePath))
            {
                TaskDialog.Show("작업 취소됨", "간격 또는 CSV 파일이 입력되지 않았습니다.");
                return Result.Cancelled;
            }

            spacingMM = wpfWindow.SpacingMM.Value;
            csvFilePath = wpfWindow.CsvFilePath;
            SaveSettings(spacingMM, csvFilePath);

            double spacingFT = spacingMM / 304.8;
            double toleranceFT = 2.5 / 304.8;
            double step5mm = 5.0 / 304.8;

            try
            {
                IList<Reference> pipeRefs = uidoc.Selection.PickObjects(ObjectType.Element, new PipeSelectionFilter(), "정렬할 파이프들을 선택하세요.");
                if (pipeRefs.Count < 2)
                {
                    TaskDialog.Show("오류", "최소 두 개 이상의 파이프를 선택해야 합니다.");
                    return Result.Failed;
                }

                Reference refPipeRef = uidoc.Selection.PickObject(ObjectType.Element, new PipeSelectionFilter(), "기준 파이프를 선택하세요.");
                Pipe refPipe = doc.GetElement(refPipeRef.ElementId) as Pipe;
                if (refPipe == null)
                {
                    message = "기준 파이프가 유효하지 않습니다.";
                    return Result.Failed;
                }

                bool isMetric = GetPipeSizeUnit(doc);
                List<PipeAlignData> alignList = new List<PipeAlignData>();

                using (Transaction t = new Transaction(doc, "절연 두께 설정 및 정렬"))
                {
                    t.Start();

                    foreach (Reference r in pipeRefs)
                    {
                        Pipe pipe = doc.GetElement(r.ElementId) as Pipe;
                        if (pipe == null) continue;

                        string bmUnit = NormalizeText(GetParameterValue(pipe, "BM Unit")).ToLower();
                        string bmZone = NormalizeText(GetParameterValue(pipe, "BM Zone")).ToLower();
                        if (string.IsNullOrWhiteSpace(bmUnit)) bmUnit = "Default";
                        if (string.IsNullOrWhiteSpace(bmZone)) bmZone = "Default";

                        string systemType = NormalizeText(pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsValueString()?.Trim() ?? "N/A");
                        string diameterRaw = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.AsValueString()?.Trim() ?? "0";
                        string diameterStr = Regex.Replace(diameterRaw, @"[^0-9.]", "");
                        double.TryParse(diameterStr, out double diameter);

                        string resultStr = FindMatchingData(csvFilePath, bmUnit, bmZone, systemType, diameter, isMetric, out _, out _);
                        if (int.TryParse(resultStr, out int thickness))
                        {
                            SetParameterValue(pipe, "Insulation Fluid Thickness", thickness);
                        }

                        LocationCurve loc = pipe.Location as LocationCurve;
                        XYZ center = (loc.Curve.GetEndPoint(0) + loc.Curve.GetEndPoint(1)) / 2;
                        double proj = GetProjection(refPipe, center, out XYZ perpDir);
                        double radius = GetPipeRadius(pipe);

                        alignList.Add(new PipeAlignData
                        {
                            Pipe = pipe,
                            CenterProj = proj,
                            Radius = radius,
                            ThicknessMM = thickness,
                            Center = center
                        });
                    }

                    // 기준 파이프 속성
                    LocationCurve refLoc = refPipe.Location as LocationCurve;
                    XYZ refStart = refLoc.Curve.GetEndPoint(0);
                    XYZ refEnd = refLoc.Curve.GetEndPoint(1);
                    XYZ refCenter = (refStart + refEnd) / 2;
                    double refProj = GetProjection(refPipe, refCenter, out XYZ perpDirection);
                    double roundedRefProj = RoundToNearestStep(refProj, step5mm);
                    double refRadius = GetPipeRadius(refPipe);
                    int refThkMM = GetIntParam(refPipe, "Insulation Fluid Thickness");

                    var left = alignList.Where(p => p.CenterProj < refProj).OrderByDescending(p => p.CenterProj).ToList();
                    var right = alignList.Where(p => p.CenterProj > refProj).OrderBy(p => p.CenterProj).ToList();

                    AlignPipes(doc, left, roundedRefProj, refRadius, refThkMM, -1, spacingFT, toleranceFT, step5mm, perpDirection);
                    AlignPipes(doc, right, roundedRefProj, refRadius, refThkMM, 1, spacingFT, toleranceFT, step5mm, perpDirection);

                    t.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private void SaveSettings(double spacingMM, string csvPath)
        {
            string folder = Path.GetDirectoryName(SettingsFilePath);
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            File.WriteAllText(SettingsFilePath, $"{spacingMM}\n{csvPath}");
        }

        private void LoadSettings(out double spacingMM, out string csvPath)
        {
            spacingMM = 100.0;
            csvPath = string.Empty;

            if (File.Exists(SettingsFilePath))
            {
                var lines = File.ReadAllLines(SettingsFilePath);
                if (lines.Length >= 2)
                {
                    double.TryParse(lines[0], out spacingMM);
                    csvPath = lines[1];
                }
            }
        }


private void AlignPipes(Document doc, List<PipeAlignData> pipes, double startProj, double startRadius, int startThk, int direction, double spacingFT, double toleranceFT, double step5mm, XYZ perpDir)
        {
            double prevCenter = startProj;
            double prevRadius = startRadius;
            int prevThk = startThk;

            foreach (var pipe in pipes)
            {
                double combinedGap = spacingFT + (prevThk + pipe.ThicknessMM) / 304.8;
                double centerGap = prevRadius + pipe.Radius + combinedGap;
                double target = prevCenter + direction * centerGap;
                double rounded = RoundToNearestStep(target, step5mm);
                double actualGap = Math.Abs(rounded - prevCenter) - prevRadius - pipe.Radius;

                if (Math.Abs(actualGap - combinedGap) <= toleranceFT)
                {
                    XYZ moveVec = perpDir.Multiply(rounded - pipe.CenterProj);
                    ElementTransformUtils.MoveElement(doc, pipe.Pipe.Id, moveVec);
                    pipe.CenterProj = rounded;
                    prevCenter = rounded;
                    prevRadius = pipe.Radius;
                    prevThk = pipe.ThicknessMM;
                }
            }
        }

        private string FindMatchingData(string filePath, string bmUnit, string bmZone, string systemType, double diameter, bool isMetric, out string matchedUnit, out string matchedZone)
        {
            matchedUnit = "Default";
            matchedZone = "Default";

            try
            {
                var lines = File.ReadAllLines(filePath);
                foreach (var line in lines)
                {
                    var values = line.Split(',');
                    if (values.Length < 8) continue;

                    string csvBMUnit = NormalizeText(values[0].Trim()).ToLower();
                    string csvBMZone = NormalizeText(values[1].Trim()).ToLower();
                    if (string.IsNullOrWhiteSpace(csvBMUnit)) csvBMUnit = "default";
                    if (string.IsNullOrWhiteSpace(csvBMZone)) csvBMZone = "default";

                    string csvSystemType = NormalizeText(values[2].Trim());
                    int fromIdx = isMetric ? 3 : 5;
                    int toIdx = isMetric ? 4 : 6;

                    if (double.TryParse(values[fromIdx], out double fromVal) && double.TryParse(values[toIdx], out double toVal))
                    {
                        string[] allowedTypes = csvSystemType.Split('/');

                        if (csvBMUnit == bmUnit &&
                            csvBMZone == bmZone &&
                            allowedTypes.Any(allowed => allowed.Trim().Equals(systemType.Trim(), StringComparison.OrdinalIgnoreCase)) &&
                            diameter > fromVal && diameter <= toVal)
                        {
                            matchedUnit = csvBMUnit;
                            matchedZone = csvBMZone;
                            return values[7].Trim();
                        }
                    }
                }
            }
            catch { }

            return "0";
        }


        private double GetProjection(Pipe refPipe, XYZ point, out XYZ perpDir)
        {
            LocationCurve loc = refPipe.Location as LocationCurve;
            XYZ dir = (loc.Curve.GetEndPoint(1) - loc.Curve.GetEndPoint(0)).Normalize();
            perpDir = new XYZ(-dir.Y, dir.X, 0).Normalize();
            return perpDir.DotProduct(point);
        }

        private string NormalizeText(string input) => input?.Normalize(NormalizationForm.FormC) ?? string.Empty;

        private double GetPipeRadius(Pipe pipe)
        {
            var param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
            return (param != null && param.HasValue) ? param.AsDouble() / 2.0 : 0.0;
        }

        private int GetIntParam(Element elem, string paramName)
        {
            var param = elem.LookupParameter(paramName);
            return (param != null && param.StorageType == StorageType.Integer) ? param.AsInteger() : 0;
        }

        private string GetParameterValue(Element element, string paramName)
        {
            var param = element.LookupParameter(paramName);
            return (param == null) ? "N/A" : param.AsString()?.Trim() ?? "Default";
        }

        private void SetParameterValue(Element element, string paramName, int value)
        {
            var param = element.LookupParameter(paramName);
            if (param != null && param.StorageType == StorageType.Integer)
                param.Set(value);
        }

        private bool GetPipeSizeUnit(Document doc)
        {
            Units projectUnits = doc.GetUnits();
            FormatOptions formatOptions = projectUnits.GetFormatOptions(SpecTypeId.PipeSize);
            return formatOptions.GetUnitTypeId() == UnitTypeId.Millimeters;
        }

        private double RoundToNearestStep(double value, double step)
        {
            return Math.Round(value / step) * step;
        }

        private class PipeAlignData
        {
            public Pipe Pipe;
            public double CenterProj;
            public double Radius;
            public int ThicknessMM;
            public XYZ Center;
        }

        public class PipeSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is Pipe;
            public bool AllowReference(Reference reference, XYZ position) => false;
        }
    }
}