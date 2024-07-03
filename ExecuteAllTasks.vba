Sub ExecuteAllTasks()
    Dim FolderPath As String
    Dim FileName As String
    Dim TemplateFilePath As String
    Dim wbSource As Workbook
    Dim wbTarget As Workbook
    Dim ws As Worksheet

    ' 템플릿 파일 선택 대화 상자를 열기
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "템플릿 파일을 선택하세요"
        .AllowMultiSelect = False
        .Filters.Add "Excel Files", "*.xlsx"
        If .Show = -1 Then
            TemplateFilePath = .SelectedItems(1)
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
            FolderPath = .SelectedItems(1) & "\"
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
        
        ' 시트 복사
        wbSource.Sheets("갑지_협력사 전체 정산 확인용").Copy After:=wbTarget.Sheets(wbTarget.Sheets.Count)
        wbSource.Sheets("을지_협력사 소속 라이더 정산 확인용").Copy After:=wbTarget.Sheets(wbTarget.Sheets.Count)
        wbSource.Sheets("관리비 및 추가배달료").Copy After:=wbTarget.Sheets(wbTarget.Sheets.Count)
        
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
        매크로4 wbTarget
        
        ' 변경 내용 저장 및 파일 닫기
        wbTarget.Close SaveChanges:=True
        
        ' 다음 파일로 이동
        FileName = Dir
    Loop

    ' 완료 메시지
    MsgBox "모든 파일에 작업이 완료되었습니다."
End Sub

Sub 매크로4(wb As Workbook)
    With wb
        ' 매크로 작업 수행
        .Sheets("Sheet1").Range("C2").Copy
        .Sheets("갑지_협력사 전체 정산 확인용").Range("D5").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet1").Range("D2").Copy
        .Sheets("갑지_협력사 전체 정산 확인용").Range("D6").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet1").Range("E2").Copy
        .Sheets("갑지_협력사 전체 정산 확인용").Range("D7").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet1").Range("F2").Copy
        .Sheets("갑지_협력사 전체 정산 확인용").Range("D8").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet1").Range("A2:B2").Copy
        .Sheets("갑지_협력사 전체 정산 확인용").Range("B14").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet1").Range("J2").Copy
        .Sheets("갑지_협력사 전체 정산 확인용").Range("D14").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet1").Range("M2").Copy
        .Sheets("갑지_협력사 전체 정산 확인용").Range("E14").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet1").Range("Q2").Copy
        .Sheets("갑지_협력사 전체 정산 확인용").Range("F14").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet1").Range("S2:V2").Copy
        .Sheets("갑지_협력사 전체 정산 확인용").Range("G14").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet1").Range("W2").Copy
        .Sheets("갑지_협력사 전체 정산 확인용").Range("K14").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet1").Range("AC2").Copy
        .Sheets("갑지_협력사 전체 정산 확인용").Range("L14").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet1").Range("AD2").Copy
        .Sheets("갑지_협력사 전체 정산 확인용").Range("M14").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("갑지_협력사 전체 정산 확인용").Range("D14").Copy
        .Sheets("갑지_협력사 전체 정산 확인용").Range("C20").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("갑지_협력사 전체 정산 확인용").Range("E14").Copy
        .Sheets("갑지_협력사 전체 정산 확인용").Range("C21").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

        ' 매크로6 작업 수행
        .Sheets("Sheet2").Range("G2:I100").Copy
        .Sheets("을지_협력사 소속 라이더 정산 확인용").Range("B16").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet2").Range("L2:L100").Copy
        .Sheets("을지_협력사 소속 라이더 정산 확인용").Range("E16").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet2").Range("O2:O100").Copy
        .Sheets("을지_협력사 소속 라이더 정산 확인용").Range("F16").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet2").Range("P2:AE100").Copy
        .Sheets("을지_협력사 소속 라이더 정산 확인용").Range("G16").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet1").Range("E2").Copy
        .Sheets("관리비 및 추가배달료").Range("B4").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet1").Range("F2").Copy
        .Sheets("관리비 및 추가배달료").Range("C4").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet1").Range("D2").Copy
        .Sheets("관리비 및 추가배달료").Range("D4").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet1").Range("C2").Copy
        .Sheets("관리비 및 추가배달료").Range("E4").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet3").Range("E2:N2").Copy
        .Sheets("관리비 및 추가배달료").Range("B9").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        .Sheets("Sheet4").Range("E2:G100").Copy
        .Sheets("관리비 및 추가배달료").Range("B14").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

        ' 시트 삭제
        Application.DisplayAlerts = False
        .Sheets("Sheet1").Delete
        .Sheets("Sheet2").Delete
        .Sheets("Sheet3").Delete
        .Sheets("Sheet4").Delete
        Application.DisplayAlerts = True

        ' 모든 시트의 커서를 A1로 이동
        For Each ws In .Worksheets
            ws.Activate
            ws.Range("A1").Select
        Next ws

        ' 첫 번째 시트를 선택
        .Worksheets(1).Activate
    End With
End Sub

