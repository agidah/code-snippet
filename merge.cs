using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ExcelFileRenamer
{
    public partial class Form1 : Form
    {
        private string selectedFolderPath;
        private string selectedMergeFolderPath;
        private string selectedReferenceFilePath;

        public Form1()
        {
            InitializeComponent();
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            using (var folderBrowser = new FolderBrowserDialog())
            {
                if (folderBrowser.ShowDialog() == DialogResult.OK)
                {
                    selectedFolderPath = folderBrowser.SelectedPath;
                    LoadExcelFiles();
                }
            }
        }

        private void LoadExcelFiles()
        {
            listBoxFiles.Items.Clear();

            var files = Directory.GetFiles(selectedFolderPath, "*.xlsx")
                                 .Select(Path.GetFileName)
                                 .ToList();

            listBoxFiles.Items.AddRange(files.ToArray());
        }

        private void btnRenameFiles_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFolderPath))
            {
                MessageBox.Show("폴더를 먼저 선택해주세요.");
                return;
            }

            var files = Directory.GetFiles(selectedFolderPath, "*.xlsx");

            foreach (var filePath in files)
            {
                var fileName = Path.GetFileName(filePath);
                var newFileName = RenameFile(fileName);

                if (newFileName != fileName)
                {
                    var newFilePath = Path.Combine(selectedFolderPath, newFileName);
                    File.Move(filePath, newFilePath);
                }
            }

            MessageBox.Show("파일명이 성공적으로 변경되었습니다.");
            LoadExcelFiles();
        }

        private string RenameFile(string fileName)
        {
            // 확장자 분리
            var extension = Path.GetExtension(fileName);
            var nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);

            // + 표시 제거
            var newName = nameWithoutExtension.Replace("+", "");

            // - 표시 이후 부분 제거
            var indexOfDash = newName.IndexOf('-');
            if (indexOfDash >= 0)
            {
                newName = newName.Substring(0, indexOfDash);
            }

            // 확장자를 유지한 채로 파일명 반환
            return newName.Trim() + extension;
        }

        private void btnSelectFolderForMerge_Click(object sender, EventArgs e)
        {
            using (var folderBrowser = new FolderBrowserDialog())
            {
                if (folderBrowser.ShowDialog() == DialogResult.OK)
                {
                    selectedMergeFolderPath = folderBrowser.SelectedPath;
                    LoadMergeFiles();
                }
            }
        }

        private void LoadMergeFiles()
        {
            listBoxMergeFiles.Items.Clear();

            var files = Directory.GetFiles(selectedMergeFolderPath, "*.xlsx")
                                 .Select(Path.GetFileName)
                                 .ToList();

            listBoxMergeFiles.Items.AddRange(files.ToArray());
        }

        private void btnSelectReferenceFile_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedReferenceFilePath = openFileDialog.FileName;
                    txtReferenceFile.Text = selectedReferenceFilePath;
                }
            }
        }

        private void btnMergeFiles_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(selectedMergeFolderPath))
            {
                MessageBox.Show("폴더를 먼저 선택해주세요.");
                return;
            }

            if (string.IsNullOrEmpty(selectedReferenceFilePath))
            {
                MessageBox.Show("참조 파일을 선택해주세요.");
                return;
            }

            var password = txtPassword.Text;  // 사용자로부터 입력받은 암호

            var referenceData = LoadReferenceData(selectedReferenceFilePath);
            var files = Directory.GetFiles(selectedMergeFolderPath, "*.xlsx");
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // 헤더 추가
                worksheet.Cells[1, 1].Value = "라이더로그인ID";
                worksheet.Cells[1, 2].Value = "항목명";
                worksheet.Cells[1, 3].Value = "설명";
                worksheet.Cells[1, 4].Value = "금액";
                worksheet.Cells[1, 5].Value = "적용일자(yyyy-MM-dd)";

                int currentRow = 2;

                foreach (var filePath in files)
                {
                    using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite)) // 읽기/쓰기 모드로 스트림 열기
                    {
                        using (var excelPackage = new ExcelPackage(stream, password))
                        {
                            var sheet = excelPackage.Workbook.Worksheets[0]; // 첫 번째 시트
                            int rowCount = sheet.Dimension.Rows;

                            for (int row = 2; row <= rowCount; row++) // 첫 행 제외
                            {
                                var colG = sheet.Cells[row, 7].Text; // G열 값 확인
                                if (colG == "달성")
                                {
                                    var colB = sheet.Cells[row, 2].Text;
                                    var colD = sheet.Cells[row, 4].Text;
                                    var colC = Path.GetFileNameWithoutExtension(filePath).Replace(" ", ""); // 파일명 추가 (확장자 제외, 공백 제거)
                                    worksheet.Cells[currentRow, 1].Value = colB;
                                    worksheet.Cells[currentRow, 2].Value = colD;
                                    worksheet.Cells[currentRow, 3].Value = colC;

                                    if (referenceData.ContainsKey(colC))
                                    {
                                        worksheet.Cells[currentRow, 4].Value = referenceData[colC].Item1;
                                        worksheet.Cells[currentRow, 5].Value = referenceData[colC].Item2;
                                    }

                                    currentRow++;
                                }
                            }
                        }
                    }
                }

                var outputFilePath = Path.Combine(selectedMergeFolderPath, "정산 업로드.xlsx");
                package.SaveAs(new FileInfo(outputFilePath));
            }

            MessageBox.Show("파일이 성공적으로 합쳐졌습니다.");
        }

        private Dictionary<string, Tuple<string, string>> LoadReferenceData(string referenceFilePath)
        {
            var referenceData = new Dictionary<string, Tuple<string, string>>();

            using (var stream = new FileStream(referenceFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var excelPackage = new ExcelPackage(stream))
                {
                    var sheet = excelPackage.Workbook.Worksheets[0]; // 첫 번째 시트
                    int rowCount = sheet.Dimension.Rows;

                    for (int row = 1; row <= rowCount; row++)
                    {
                        var key = sheet.Cells[row, 1].Text.Replace(" ", ""); // A열, 공백 제거
                        var value1 = sheet.Cells[row, 2].Text; // B열
                        var value2 = sheet.Cells[row, 3].Text; // C열
                        if (!string.IsNullOrEmpty(key) && !referenceData.ContainsKey(key))
                        {
                            referenceData.Add(key, new Tuple<string, string>(value1, value2));
                        }
                    }
                }
            }

            return referenceData;
        }
    }
}
