using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace PipeInsulationInput
{
    public partial class InsulationSettingsWindow : Window
    {
        private static readonly string AppDataDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MTMDigCon");
        private static readonly string SettingsFile = Path.Combine(AppDataDir, "user_settings.json");
        private static UserSettings LastSettings = null;

        public bool UseBMArea => chkBMArea.IsChecked == true;
        public bool UseBMUnit => chkBMUnit.IsChecked == true;
        public bool UseBMZone => chkBMZone.IsChecked == true;
        public bool UseFluid => chkFluid.IsChecked == true;
        public bool UseSystemType => chkSystemType.IsChecked == true;
        public string CsvFilePath { get; private set; }

        public string BMAreaParamName => cmbBMAreaParam.Text;
        public string BMUnitParamName => cmbBMUnitParam.Text;
        public string BMZoneParamName => cmbBMZoneParam.Text;
        public string FluidParamName => cmbFluidParam.Text;
        // SystemTypeParamName 없음

        public InsulationSettingsWindow(Document doc)
        {
            InitializeComponent();

            // CSV, 체크박스, 파라미터 기본값/복원
            var settings = LoadLastSettings();
            if (settings != null)
            {
                CsvFilePath = settings.LastCsvFilePath ?? "";
                txtCsvPath.Text = CsvFilePath;
                chkBMArea.IsChecked = settings.UseBMArea;
                chkBMUnit.IsChecked = settings.UseBMUnit;
                chkBMZone.IsChecked = settings.UseBMZone;
                chkFluid.IsChecked = settings.UseFluid;
                chkSystemType.IsChecked = settings.UseSystemType;
                cmbBMAreaParam.Text = string.IsNullOrWhiteSpace(settings.BMAreaParamName) ? "BM Area" : settings.BMAreaParamName;
                cmbBMUnitParam.Text = string.IsNullOrWhiteSpace(settings.BMUnitParamName) ? "BM Unit" : settings.BMUnitParamName;
                cmbBMZoneParam.Text = string.IsNullOrWhiteSpace(settings.BMZoneParamName) ? "BM Zone" : settings.BMZoneParamName;
                cmbFluidParam.Text = string.IsNullOrWhiteSpace(settings.FluidParamName) ? "Fluid" : settings.FluidParamName;
            }
            else
            {
                // ✅ 최초 실행 시 기본값 설정
                CsvFilePath = "";
                txtCsvPath.Text = "";
                // 콤보박스 초기 텍스트
                cmbBMAreaParam.Text = "BM Area";
                cmbBMUnitParam.Text = "BM Unit";
                cmbBMZoneParam.Text = "BM Zone";
                cmbFluidParam.Text = "Fluid";

                // ✅ 최초 실행 시 체크 여부 기본값 (여기서 설정 가능)
                chkBMArea.IsChecked = false;
                chkBMUnit.IsChecked = false;
                chkBMZone.IsChecked = true;
                chkFluid.IsChecked = false;
                chkSystemType.IsChecked = true;

            }

            //selectCSV.Click += SelectCSV_Click;
            Loaded += (s, e) => { LoadParameterNames(doc); };

            // 체크박스 변경 시 ComboBox Enable/Disable 연동
            chkBMArea.Checked += (s, e) => cmbBMAreaParam.IsEnabled = true;
            chkBMArea.Unchecked += (s, e) => cmbBMAreaParam.IsEnabled = false;
            chkBMUnit.Checked += (s, e) => cmbBMUnitParam.IsEnabled = true;
            chkBMUnit.Unchecked += (s, e) => cmbBMUnitParam.IsEnabled = false;
            chkBMZone.Checked += (s, e) => cmbBMZoneParam.IsEnabled = true;
            chkBMZone.Unchecked += (s, e) => cmbBMZoneParam.IsEnabled = false;
            chkFluid.Checked += (s, e) => cmbFluidParam.IsEnabled = true;
            chkFluid.Unchecked += (s, e) => cmbFluidParam.IsEnabled = false;
            // System Type 콤보박스 없음

            // 최초 상태 반영
            cmbBMAreaParam.IsEnabled = chkBMArea.IsChecked == true;
            cmbBMUnitParam.IsEnabled = chkBMUnit.IsChecked == true;
            cmbBMZoneParam.IsEnabled = chkBMZone.IsChecked == true;
            cmbFluidParam.IsEnabled = chkFluid.IsChecked == true;
            // System Type 콤보박스 없음
        }

        private void SelectCSV_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = "CSV 파일 (*.csv)|*.csv"
            };
            if (!string.IsNullOrWhiteSpace(CsvFilePath))
            {
                try { dlg.InitialDirectory = Path.GetDirectoryName(CsvFilePath); } catch { }
            }
            if (dlg.ShowDialog() == true)
            {
                CsvFilePath = dlg.FileName;
                txtCsvPath.Text = CsvFilePath;
                SaveLastSettings();
            }
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(CsvFilePath))
            {
                MessageBox.Show("CSV 파일을 선택하세요.", "오류", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            SaveLastSettings();
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void LoadParameterNames(Document doc)
        {
            var paramNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var pipes = new FilteredElementCollector(doc).OfClass(typeof(Pipe)).WhereElementIsNotElementType();
            var fittings = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType();
            var accessories = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType();

            foreach (var elem in pipes.Concat(fittings).Concat(accessories))
            {
                foreach (Parameter param in elem.Parameters)
                    if (!string.IsNullOrWhiteSpace(param.Definition.Name))
                        paramNames.Add(param.Definition.Name);
            }
            var sortedParamNames = paramNames.OrderBy(n => n).ToList();
            cmbBMAreaParam.ItemsSource = sortedParamNames;
            cmbBMUnitParam.ItemsSource = sortedParamNames;
            cmbBMZoneParam.ItemsSource = sortedParamNames;
            cmbFluidParam.ItemsSource = sortedParamNames;
            // System Type 콤보박스 없음
        }

        private static UserSettings LoadLastSettings()
        {
            try
            {
                if (LastSettings != null)
                    return LastSettings;
                if (File.Exists(SettingsFile))
                {
                    var json = File.ReadAllText(SettingsFile);
                    var data = JsonConvert.DeserializeObject<UserSettings>(json);
                    if (data != null)
                    {
                        LastSettings = data;
                        return data;
                    }
                }
            }
            catch { }
            return null;
        }

        private void SaveLastSettings()
        {
            try
            {
                if (!Directory.Exists(AppDataDir))
                    Directory.CreateDirectory(AppDataDir);

                var data = new UserSettings
                {
                    LastCsvFilePath = CsvFilePath,
                    UseBMArea = chkBMArea.IsChecked == true,
                    UseBMUnit = chkBMUnit.IsChecked == true,
                    UseBMZone = chkBMZone.IsChecked == true,
                    UseFluid = chkFluid.IsChecked == true,
                    UseSystemType = chkSystemType.IsChecked == true,
                    BMAreaParamName = cmbBMAreaParam.Text,
                    BMUnitParamName = cmbBMUnitParam.Text,
                    BMZoneParamName = cmbBMZoneParam.Text,
                    FluidParamName = cmbFluidParam.Text
                    // System Type 매핑 파라미터 저장/복원 없음
                };
                File.WriteAllText(SettingsFile, JsonConvert.SerializeObject(data));
                LastSettings = data;
            }
            catch { }
        }

        private class UserSettings
        {
            public string LastCsvFilePath { get; set; }
            public bool UseBMArea { get; set; }
            public bool UseBMUnit { get; set; }
            public bool UseBMZone { get; set; }
            public bool UseFluid { get; set; }
            public bool UseSystemType { get; set; }
            public string BMAreaParamName { get; set; }
            public string BMUnitParamName { get; set; }
            public string BMZoneParamName { get; set; }
            public string FluidParamName { get; set; }
            public string SystemTypeParamName { get; set; }
            // public string SystemTypeParamName { get; set; } // 삭제됨
        }
    }
}