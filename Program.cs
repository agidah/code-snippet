//Program.cs

using ccal;
using System;
using System.Windows.Forms;

namespace ExcelRowSplitter
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}




// Designer
namespace ExcelRowSplitter
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.btnAttachFile = new System.Windows.Forms.Button();
            this.lblFilePath = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // btnAttachFile
            // 
            this.btnAttachFile.Location = new System.Drawing.Point(12, 12);
            this.btnAttachFile.Name = "btnAttachFile";
            this.btnAttachFile.Size = new System.Drawing.Size(125, 23);
            this.btnAttachFile.TabIndex = 0;
            this.btnAttachFile.Text = "데이터파일 첨부";
            this.btnAttachFile.UseVisualStyleBackColor = true;
            this.btnAttachFile.Click += new System.EventHandler(this.btnAttachFile_Click);
            // 
            // lblFilePath
            // 
            this.lblFilePath.AutoSize = true;
            this.lblFilePath.Location = new System.Drawing.Point(143, 17);
            this.lblFilePath.Name = "lblFilePath";
            this.lblFilePath.Size = new System.Drawing.Size(0, 13);
            this.lblFilePath.TabIndex = 1;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 50);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(760, 23);
            this.progressBar.TabIndex = 2;
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblFilePath);
            this.Controls.Add(this.btnAttachFile);
            this.Name = "Form1";
            this.Text = "정산서 발급기";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.Button btnAttachFile;
        private System.Windows.Forms.Label lblFilePath;
        private System.Windows.Forms.ProgressBar progressBar;
    }
}



// form

using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelRowSplitter
{
    public partial class Form1 : Form
    {
        private string outputDirectory;
        private string[] headerRow;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnAttachFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx",
                Title = "전체 데이터 엑셀파일 선택하세요"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                lblFilePath.Text = filePath;

                using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "정산서를 배포할 폴더 선택하세요";
                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        outputDirectory = folderDialog.SelectedPath;
                        ProcessExcelFile(filePath);
                    }
                    else
                    {
                        MessageBox.Show("No folder selected. Operation cancelled.");
                    }
                }
            }
        }

        private void ProcessExcelFile(string filePath)
        {
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(fs);

                    if (workbook.NumberOfSheets < 4)
                    {
                        MessageBox.Show("The Excel file does not contain at least 4 sheets.");
                        return;
                    }

                    ISheet sheet1 = workbook.GetSheetAt(0);
                    ISheet sheet2 = workbook.GetSheetAt(1);
                    ISheet sheet3 = workbook.GetSheetAt(2);
                    ISheet sheet4 = workbook.GetSheetAt(3);

                    int rowCount = sheet1.PhysicalNumberOfRows;

                    IRow header = sheet1.GetRow(0);
                    headerRow = new string[32];
                    for (int col = 0; col < 32; col++)
                    {
                        headerRow[col] = header.GetCell(col)?.ToString();
                    }

                    progressBar.Maximum = rowCount - 1;
                    progressBar.Value = 0;

                    for (int row = 1; row < rowCount; row++)
                    {
                        IRow currentRow = sheet1.GetRow(row);
                        string[] rowData = new string[32];
                        for (int col = 0; col < 32; col++)
                        {
                            rowData[col] = currentRow.GetCell(col)?.ToString();
                        }

                        string fileName = $"{rowData[0]}_{rowData[1]}_{rowData[2]}_{rowData[3]}_{rowData[4]}_{rowData[5]}.xlsx";
                        SaveRowToNewExcelFile(rowData, fileName, sheet1, sheet2, sheet3, sheet4);

                        progressBar.Value++;
                    }
                }

                MessageBox.Show("정산서 작성이 완료되었습니다!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

        private void SaveRowToNewExcelFile(string[] rowData, string fileName, ISheet sheet1, ISheet sheet2, ISheet sheet3, ISheet sheet4)
        {
            IWorkbook newWorkbook = new XSSFWorkbook();
            ISheet newSheet1 = newWorkbook.CreateSheet("Sheet1");
            ISheet newSheet2 = newWorkbook.CreateSheet("Sheet2");
            ISheet newSheet3 = newWorkbook.CreateSheet("Sheet3");
            ISheet newSheet4 = newWorkbook.CreateSheet("Sheet4");

            string sheet1C2Value = rowData[2]; // Assuming C2 value in Sheet1 corresponds to rowData[2]

            // Sheet1: 제목 행 추가
            IRow headerRowInNewFile = newSheet1.CreateRow(0);
            for (int col = 0; col < headerRow.Length; col++)
            {
                headerRowInNewFile.CreateCell(col).SetCellValue(headerRow[col]);
            }

            // Sheet1: 데이터 행 추가
            IRow newRow = newSheet1.CreateRow(1);
            for (int col = 0; col < rowData.Length; col++)
            {
                newRow.CreateCell(col).SetCellValue(rowData[col]);
            }

            // Sheet2, Sheet3, Sheet4: 데이터 복사 및 비교 후 삭제
            CopyAndFilterSheet(sheet2, newSheet2, sheet1C2Value, 2); // Compare with column C
            CopyAndFilterSheet(sheet3, newSheet3, sheet1C2Value, 3); // Compare with column D
            CopyAndFilterSheet(sheet4, newSheet4, sheet1C2Value, 3); // Compare with column E

            // 빈 행 제거
            RemoveEmptyRows(newSheet1);
            RemoveEmptyRows(newSheet2);
            RemoveEmptyRows(newSheet3);
            RemoveEmptyRows(newSheet4);

            string savePath = Path.Combine(outputDirectory, fileName);
            using (FileStream fs = new FileStream(savePath, FileMode.Create, FileAccess.Write))
            {
                newWorkbook.Write(fs);
            }
        }

        private void CopyAndFilterSheet(ISheet sourceSheet, ISheet targetSheet, string compareValue, int compareColumnIndex)
        {
            IRow headerRow = sourceSheet.GetRow(0);
            IRow newHeaderRow = targetSheet.CreateRow(0);

            for (int col = 0; col < headerRow.LastCellNum; col++)
            {
                newHeaderRow.CreateCell(col).SetCellValue(headerRow.GetCell(col).ToString());
            }

            int targetRowIndex = 1;

            for (int i = 1; i <= sourceSheet.LastRowNum; i++)
            {
                IRow sourceRow = sourceSheet.GetRow(i);
                if (sourceRow == null) continue;

                ICell compareCell = sourceRow.GetCell(compareColumnIndex);
                if (compareCell != null && compareCell.ToString() == compareValue)
                {
                    IRow targetRow = targetSheet.CreateRow(targetRowIndex++);
                    for (int j = 0; j < sourceRow.LastCellNum; j++)
                    {
                        ICell sourceCell = sourceRow.GetCell(j);
                        ICell targetCell = targetRow.CreateCell(j);
                        if (sourceCell != null)
                        {
                            targetCell.SetCellValue(sourceCell.ToString());
                        }
                    }
                }
            }
        }

        private void RemoveEmptyRows(ISheet sheet)
        {
            for (int i = sheet.LastRowNum; i > 0; i--)
            {
                IRow row = sheet.GetRow(i);
                if (row == null || row.Cells.All(d => d.CellType == CellType.Blank))
                {
                    sheet.RemoveRow(row);
                }
            }
        }
    }
}

