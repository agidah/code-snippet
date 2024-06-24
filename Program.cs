//Program.cs

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
            this.SuspendLayout();
            // 
            // btnAttachFile
            // 
            this.btnAttachFile.Location = new System.Drawing.Point(12, 12);
            this.btnAttachFile.Name = "btnAttachFile";
            this.btnAttachFile.Size = new System.Drawing.Size(125, 23);
            this.btnAttachFile.TabIndex = 0;
            this.btnAttachFile.Text = "Attach Excel File";
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
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.lblFilePath);
            this.Controls.Add(this.btnAttachFile);
            this.Name = "Form1";
            this.Text = "Excel Row Splitter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.Button btnAttachFile;
        private System.Windows.Forms.Label lblFilePath;
    }
}


// form

using System;
using System.IO;
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
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                lblFilePath.Text = filePath;

                // 사용자에게 추출할 폴더를 선택하게 함
                using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "Select the folder to save the extracted files";
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

                    // 시트가 하나 이상 존재하는지 확인
                    if (workbook.NumberOfSheets == 0)
                    {
                        MessageBox.Show("The Excel file does not contain any sheets.");
                        return;
                    }

                    ISheet sheet = workbook.GetSheetAt(0);
                    int rowCount = sheet.PhysicalNumberOfRows;

                    // 제목 행 가져오기
                    IRow header = sheet.GetRow(0);
                    headerRow = new string[32]; // A에서 AD열까지 총 32열
                    for (int col = 0; col < 32; col++) // A열은 0, AD열은 31
                    {
                        headerRow[col] = header.GetCell(col)?.ToString();
                    }

                    for (int row = 1; row < rowCount; row++) // Assuming first row is header
                    {
                        IRow currentRow = sheet.GetRow(row);
                        string[] rowData = new string[32]; // A에서 AD열까지 총 32열
                        for (int col = 0; col < 32; col++) // A열은 0, AD열은 31
                        {
                            rowData[col] = currentRow.GetCell(col)?.ToString();
                        }

                        string fileName = $"{rowData[0]}_{rowData[1]}_{rowData[2]}_{rowData[3]}_{rowData[4]}_{rowData[5]}.xlsx";
                        SaveRowToNewExcelFile(rowData, fileName);
                    }
                }

                MessageBox.Show("Process completed successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

        private void SaveRowToNewExcelFile(string[] rowData, string fileName)
        {
            IWorkbook newWorkbook = new XSSFWorkbook();
            ISheet newSheet = newWorkbook.CreateSheet("Sheet1");

            // 제목 행 추가
            IRow headerRowInNewFile = newSheet.CreateRow(0);
            for (int col = 0; col < headerRow.Length; col++)
            {
                headerRowInNewFile.CreateCell(col).SetCellValue(headerRow[col]);
            }

            // 데이터 행 추가
            IRow newRow = newSheet.CreateRow(1);
            for (int col = 0; col < rowData.Length; col++)
            {
                newRow.CreateCell(col).SetCellValue(rowData[col]);
            }

            string savePath = Path.Combine(outputDirectory, fileName);
            using (FileStream fs = new FileStream(savePath, FileMode.Create, FileAccess.Write))
            {
                newWorkbook.Write(fs);
            }
        }
    }
}
