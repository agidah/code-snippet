using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelRowSplitter
{
    public partial class Settlement : Form
    {
        private string outputDirectory; // 저장할 폴더 경로
        private string[] headerRow; // 헤더 행 데이터 저장

        public Settlement()
        {
            InitializeComponent(); // 폼 초기화
        }

        // 파일 첨부 버튼 클릭 이벤트 핸들러
        private void btnAttachFile_Click(object sender, EventArgs e)
        {
            // 파일 열기 대화상자 설정
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx",
                Title = "전체 데이터 엑셀파일 선택하세요"
            };

            // 파일 선택 확인
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName; // 파일 경로 저장
                lblFilePath.Text = filePath; // 파일 경로 라벨 업데이트

                // 폴더 선택 대화상자 설정
                using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "정산서를 배포할 폴더 선택하세요";
                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        outputDirectory = folderDialog.SelectedPath; // 폴더 경로 저장
                        ProcessExcelFile(filePath); // 엑셀 파일 처리
                    }
                    else
                    {
                        MessageBox.Show("폴더가 선택되지 않았습니다. 작업이 취소되었습니다.");
                    }
                }
            }
        }

        // 엑셀 파일 처리 메서드
        private void ProcessExcelFile(string filePath)
        {
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(fs); // 엑셀 파일 로드

                    // 시트가 4개 이상 있는지 확인
                    if (workbook.NumberOfSheets < 4)
                    {
                        MessageBox.Show("엑셀 파일에 시트가 4개 이상 존재하지 않습니다.");
                        return;
                    }

                    // 각 시트 로드
                    ISheet sheet1 = workbook.GetSheetAt(0);
                    ISheet sheet2 = workbook.GetSheetAt(1);
                    ISheet sheet3 = workbook.GetSheetAt(2);
                    ISheet sheet4 = workbook.GetSheetAt(3);

                    int rowCount = sheet1.PhysicalNumberOfRows; // 시트1의 총 행 수

                    // 시트1의 헤더 행 데이터 저장
                    IRow header = sheet1.GetRow(0);
                    headerRow = new string[32];
                    for (int col = 0; col < 32; col++)
                    {
                        headerRow[col] = header.GetCell(col)?.ToString();
                    }

                    // 진행바 설정
                    progressBar.Maximum = rowCount - 1;
                    progressBar.Value = 0;

                    // 데이터 행 처리
                    for (int row = 1; row < rowCount; row++)
                    {
                        IRow currentRow = sheet1.GetRow(row);
                        string[] rowData = new string[32];
                        for (int col = 0; col < 32; col++)
                        {
                            rowData[col] = currentRow.GetCell(col)?.ToString();
                        }

                        // 파일명 생성 및 특수문자 제거
                        string fileName = CleanFileName($"{rowData[0]}~{rowData[1]}_{rowData[5]}_{rowData[3]}_{rowData[2]}.xlsx");
                        // 새로운 엑셀 파일로 데이터 저장
                        SaveRowToNewExcelFile(rowData, fileName, sheet1, sheet2, sheet3, sheet4);

                        progressBar.Value++; // 진행바 업데이트
                    }
                }

                MessageBox.Show("정산서 데이터추출 완료!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"오류가 발생했습니다: {ex.Message}");
            }
        }

        // 파일명에서 특수문자 제거
        private string CleanFileName(string fileName)
        {
            return Regex.Replace(fileName, @"[\\/:*?""<>|]", string.Empty);
        }

        // 새로운 엑셀 파일로 데이터 저장
        private void SaveRowToNewExcelFile(string[] rowData, string fileName, ISheet sheet1, ISheet sheet2, ISheet sheet3, ISheet sheet4)
        {
            IWorkbook newWorkbook = new XSSFWorkbook();
            ISheet newSheet1 = newWorkbook.CreateSheet("Sheet1");
            ISheet newSheet2 = newWorkbook.CreateSheet("Sheet2");
            ISheet newSheet3 = newWorkbook.CreateSheet("Sheet3");
            ISheet newSheet4 = newWorkbook.CreateSheet("Sheet4");

            // 시트1의 C2 셀 값 저장
            string sheet1C2Value = rowData[2]; // Assuming C2 value in Sheet1 corresponds to rowData[2]

            // 시트1: 헤더 행 추가
            IRow headerRowInNewFile = newSheet1.CreateRow(0);
            for (int col = 0; col < headerRow.Length; col++)
            {
                headerRowInNewFile.CreateCell(col).SetCellValue(headerRow[col]);
            }

            // 시트1: 데이터 행 추가
            IRow newRow = newSheet1.CreateRow(1);
            for (int col = 0; col < rowData.Length; col++)
            {
                newRow.CreateCell(col).SetCellValue(rowData[col]);
            }

            // 시트2, 시트3, 시트4: 데이터 복사 및 비교 후 필터링
            CopyAndFilterSheet(sheet2, newSheet2, sheet1C2Value, 2); // 시트2의 열C와 비교
            CopyAndFilterSheet(sheet3, newSheet3, sheet1C2Value, 3); // 시트3의 열D와 비교
            CopyAndFilterSheet(sheet4, newSheet4, sheet1C2Value, 3); // 시트4의 열E와 비교

            // 특정 열 제거
            RemoveColumns(newSheet1, new[] { 6, 7, 8, 10, 11, 13, 14, 15, 17, 23, 24, 25, 26, 27 });
            RemoveColumns(newSheet2, new[] { 9, 10, 12, 13 });

            // 범위 처리
            ProcessRange(newSheet1, "G2", "AD");
            ProcessRange(newSheet2, "I2", "AD");
            ProcessRange(newSheet3, "E2", "N");
            ProcessRange(newSheet4, "G2", "G");

            // 빈 행 제거
            RemoveEmptyRows(newSheet1);
            RemoveEmptyRows(newSheet2);
            RemoveEmptyRows(newSheet3);
            RemoveEmptyRows(newSheet4);

            // 파일 저장
            string savePath = Path.Combine(outputDirectory, fileName);
            using (FileStream fs = new FileStream(savePath, FileMode.Create, FileAccess.Write))
            {
                newWorkbook.Write(fs);
            }
        }

        // 시트 복사 및 필터링
        private void CopyAndFilterSheet(ISheet sourceSheet, ISheet targetSheet, string compareValue, int compareColumnIndex)
        {
            // 헤더 행 복사
            IRow headerRow = sourceSheet.GetRow(0);
            IRow newHeaderRow = targetSheet.CreateRow(0);

            for (int col = 0; col < headerRow.LastCellNum; col++)
            {
                newHeaderRow.CreateCell(col).SetCellValue(headerRow.GetCell(col).ToString());
            }

            int targetRowIndex = 1;

            // 데이터 행 필터링 및 복사
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

        // 빈 행 제거
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

        // 범위 처리 (숫자 형식 변경)
        private void ProcessRange(ISheet sheet, string startCellAddress, string endColumnLetter)
        {
            int startRow = CellReference.ConvertCellReference(startCellAddress).Row;
            int startColumn = CellReference.ConvertCellReference(startCellAddress).Col;

            // 사용된 범위의 마지막 행 찾기
            int lastRow = sheet.LastRowNum;

            // 마지막 열 찾기
            int endColumn = CellReference.ConvertCellReference(endColumnLetter + "1").Col;

            for (int row = startRow; row <= lastRow; row++)
            {
                IRow currentRow = sheet.GetRow(row);
                if (currentRow == null) continue;

                for (int col = startColumn; col <= endColumn; col++)
                {
                    ICell cell = currentRow.GetCell(col);
                    if (cell != null && cell.CellType == CellType.String && double.TryParse(cell.StringCellValue, out double result))
                    {
                        cell.SetCellValue(result); // 값 설정
                        ICellStyle cellStyle = sheet.Workbook.CreateCellStyle();
                        IDataFormat dataFormat = sheet.Workbook.CreateDataFormat();
                        cellStyle.DataFormat = dataFormat.GetFormat("#,##0");
                        cell.CellStyle = cellStyle; // 스타일 적용 (천 단위 구분기호)
                    }
                }
            }
        }

        // 특정 열 제거
        private void RemoveColumns(ISheet sheet, int[] columnIndexes)
        {
            foreach (var columnIndex in columnIndexes.OrderByDescending(c => c))
            {
                foreach (IRow row in sheet)
                {
                    if (row.GetCell(columnIndex) != null)
                    {
                        row.RemoveCell(row.GetCell(columnIndex));
                    }
                }
            }
        }

        // 폼 로드 이벤트 핸들러 (현재는 비어 있음) 수정할 필요 없음
        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }

    // 셀 주소 변환 유틸리티 클래스
    public static class CellReference
    {
        // 셀 주소를 행과 열 인덱스로 변환
        public static (int Row, int Col) ConvertCellReference(string cellReference)
        {
            int row = 0, col = 0;
            foreach (char c in cellReference)
            {
                if (char.IsDigit(c))
                {
                    row = row * 10 + (c - '0');
                }
                else
                {
                    col = col * 26 + (char.ToUpper(c) - 'A' + 1);
                }
            }
            return (row - 1, col - 1); // NPOI는 0부터 인덱스를 사용
        }
    }
}
