using MTMDCCommon;
﻿using System;
using System.IO;
using System.Windows;
using System.Text;
using Microsoft.Win32;

namespace PipeSpacingAlign
{
    public partial class SpacingInputWindow : Window
    {
        public double? SpacingMM { get; private set; }
        public string CsvFilePath { get; private set; }

        public static double? LastSpacingMM { get; private set; }
        public static string LastCsvFilePath { get; private set; }

        public SpacingInputWindow()
        {
            InitializeComponent();
        }

        public SpacingInputWindow(double? lastSpacing, string lastCsvPath)
        {
            InitializeComponent();

            if (lastSpacing.HasValue)
            {
                SpacingTextBox.Text = lastSpacing.Value.ToString();
                LastSpacingMM = lastSpacing;
            }
            else
            {
                SpacingTextBox.Text = "100";
            }

            if (!string.IsNullOrEmpty(lastCsvPath))
            {
                CsvPathTextBox.Text = lastCsvPath;
                LastCsvFilePath = lastCsvPath;
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(SpacingTextBox.Text, out double spacingValue) && spacingValue > 0)
            {
                SpacingMM = spacingValue;
                CsvFilePath = CsvPathTextBox.Text;

                LastSpacingMM = SpacingMM;
                LastCsvFilePath = CsvFilePath;

                this.DialogResult = true;
                this.Close();
            }
            else
            {
                MessageBox.Show("유효한 간격(mm) 값을 입력하세요.", "입력 오류", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void BrowseCsvButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "CSV 파일 (*.csv)|*.csv",
                Title = "CSV 파일 선택"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                CsvPathTextBox.Text = openFileDialog.FileName;
            }
        }
    }
}