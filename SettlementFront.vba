Sub ExecuteAllTasks3()
    Dim FolderPath As String
    Dim FileName As String
    Dim TemplateFilePath As String
    Dim wbSource As Workbook
    Dim wbTarget As Workbook
    Dim ws As Worksheet

    ' 최적화를 위한 설정
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 템플릿 파일 선택 대화 상자를 열기
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "템플릿 파일을 선택하세요"
        .AllowMultiSelect = False
        .Filters.Add "Excel Files", "*.xlsx"
        If .Show = -1 Then
            TemplateFilePath = .SelectedItems(1) ' 선택한 파일 경로 저장
        Else
            Exit Sub ' 파일 선택을 취소한 경우 종료
        End If
    End With

    ' 템플릿 파일 열기
    Set wbSource = Workbooks.Open(TemplateFilePath)

    ' 폴더 선택 대화 상자를 열기
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "폴더를 선택하세요"
        .AllowMultiSelect = False
        If .Show = -1 Then
            FolderPath = .SelectedItems(1) & "\" ' 선택한 폴더 경로 저장
        Else
            wbSource.Close SaveChanges:=False
            Exit Sub ' 폴더 선택을 취소한 경우 종료
        End If
    End With

    ' 폴더 내 모든 엑셀 파일 순회 - 첫 번째 단계: 템플릿 시트 복사
    FileName = Dir(FolderPath & "*.xlsx")
    Do While FileName <> ""
        ' 현재 파일 열기
        Set wbTarget = Workbooks.Open(FolderPath & FileName)
        
        ' 템플릿 시트 복사
        CopySheets wbSource, wbTarget
        
        ' 변경 내용 저장 및 파일 닫기
        wbTarget.Close SaveChanges:=True
        
        ' 다음 파일로 이동
        FileName = Dir
    Loop

    ' 템플릿 파일 닫기
    wbSource.Close SaveChanges:=False

    ' 폴더 내 모든 엑셀 파일 순회 - 두 번째 단계: 매크로 작업 수행
    FileName = Dir(FolderPath & "*.xlsx")
    Do While FileName <> ""
        ' 현재 파일 열기
        Set wbTarget = Workbooks.Open(FolderPath & FileName)
        
        ' 매크로 작업 수행
        Macro4 wbTarget
        
        ' 변경 내용 저장 및 파일 닫기
        wbTarget.Close SaveChanges:=True
        
        ' 다음 파일로 이동
        FileName = Dir
    Loop

    ' 최적화 설정 복원
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    ' 완료 메시지
    MsgBox "모든 파일에 작업이 완료되었습니다."
End Sub

Sub CopySheets(wbSource As Workbook, wbTarget As Workbook)
    wbSource.Sheets("갑지_협력사 전체 정산 확인용").Copy After:=wbTarget.Sheets(wbTarget.Sheets.Count)
    wbSource.Sheets("을지_협력사 소속 라이더 정산 확인용").Copy After:=wbTarget.Sheets(wbTarget.Sheets.Count)
    wbSource.Sheets("관리비 및 추가배달료").Copy After:=wbTarget.Sheets(wbTarget.Sheets.Count)
End Sub

Sub Macro4(wb As Workbook)
    Dim wsSource1 As Worksheet, wsSource2 As Worksheet, wsSource3 As Worksheet, wsSource4 As Worksheet
    Dim wsTarget1 As Worksheet, wsTarget2 As Worksheet, wsTarget3 As Worksheet
    
    Set wsSource1 = wb.Sheets("Sheet1")
    Set wsSource2 = wb.Sheets("Sheet2")
    Set wsSource3 = wb.Sheets("Sheet3")
    Set wsSource4 = wb.Sheets("Sheet4")
    
    Set wsTarget1 = wb.Sheets("갑지_협력사 전체 정산 확인용")
    Set wsTarget2 = wb.Sheets("을지_협력사 소속 라이더 정산 확인용")
    Set wsTarget3 = wb.Sheets("관리비 및 추가배달료")
    
    ' 매크로 작업 수행 - 특정 셀 범위를 복사하여 붙여넣기
    wsTarget1.Range("D5").Resize(4, 1).Value = Application.Transpose(wsSource1.Range("C2:F2").Value)
    wsTarget1.Range("B14:C14").Value = wsSource1.Range("A2:B2").Value
    wsTarget1.Range("D14").Value = wsSource1.Range("J2").Value
    wsTarget1.Range("E14").Value = wsSource1.Range("M2").Value
    wsTarget1.Range("B20:D20").Value = wsSource1.Range("P2:R2").Value
    wsTarget1.Range("F14").Value = wsSource1.Range("Q2").Value
    wsTarget1.Range("G14:J14").Value = wsSource1.Range("S2:V2").Value
    wsTarget1.Range("K14").Value = wsSource1.Range("W2").Value
    wsTarget1.Range("L14").Value = wsSource1.Range("Z2").Value
    wsTarget1.Range("M14").Value = wsSource1.Range("AC2").Value
    wsTarget1.Range("N14").Value = wsSource1.Range("AD2").Value
    
    wsTarget2.Range("B16:D115").Value = wsSource2.Range("G2:I101").Value
    wsTarget2.Range("E16:E115").Value = wsSource2.Range("L2:L101").Value
    wsTarget2.Range("F16:F115").Value = wsSource2.Range("O2:O101").Value
    wsTarget2.Range("G16:U115").Value = wsSource2.Range("P2:AE101").Value
    
    wsTarget3.Range("B4:E4").Value = Application.Transpose(wsSource1.Range("C2:F2").Value)
    wsTarget3.Range("B9:K9").Value = wsSource3.Range("E2:N2").Value
    wsTarget3.Range("B14:D113").Value = wsSource4.Range("E2:G101").Value
    
    ' 숫자 형식 적용
    wsTarget1.Range("D14:N14").NumberFormat = "_ * #,##0_ ;-* #,##0_ ;-_ "
    wsTarget1.Range("B20:D20").NumberFormat = "_ * #,##0_ ;-* #,##0_ ;-_ "
    wsTarget2.Range("E16:U115").NumberFormat = "_ * #,##0_ ;-* #,##0_ ;-_ "
    
    ' 원본 시트 삭제
    Application.DisplayAlerts = False
    wsSource1.Delete
    wsSource2.Delete
    wsSource3.Delete
    wsSource4.Delete
    Application.DisplayAlerts = True
    
    ' 모든 시트의 커서를 A1로 이동
    For Each ws In wb.Worksheets
        ws.Activate
        ws.Range("A1").Select
    Next ws

    ' 첫 번째 시트를 선택
    wb.Worksheets(1).Activate
End Sub

