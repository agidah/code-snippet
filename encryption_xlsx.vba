Sub ProcessFolder()
    Dim folderPath As String
    Dim file As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fd As FileDialog
    Dim password As String

    ' 폴더 선택 대화 상자 열기
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "폴더를 선택하세요"
    
    If fd.Show = -1 Then
        folderPath = fd.SelectedItems(1) & "\"
    Else
        MsgBox "폴더를 선택하지 않았습니다. 매크로를 종료합니다."
        Exit Sub
    End If
    
    ' 폴더 내의 모든 .xlsx 파일 찾기
    file = Dir(folderPath & "*.xlsx")
    
    ' 최적화를 위해 설정 변경
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' 파일이 있는 동안 반복
    Do While file <> ""
        ' 파일 열기
        Set wb = Workbooks.Open(folderPath & file)
        
        ' 모든 시트의 커서를 A1으로 이동
        For Each ws In wb.Sheets
            ws.Activate
            ws.Range("A1").Select
        Next ws
        
        ' 첫 번째 시트를 활성화하고 D7 셀 값을 가져옴
        wb.Sheets(1).Activate
        password = wb.Sheets(1).Range("D7").Value
        
        ' 파일에 암호 설정
        wb.password = password
        
        ' 변경사항 저장 (기존 파일 덮어쓰기)
        wb.Save
        wb.Close SaveChanges:=True
        
        ' 다음 파일로 이동
        file = Dir
    Loop
    
    ' 설정 원래대로 복원
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    MsgBox "모든 파일 처리가 완료되었습니다."
End Sub

