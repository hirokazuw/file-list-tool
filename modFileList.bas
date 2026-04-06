Attribute VB_Name = "modFileList"
Option Explicit

'=============================================================================
' ファイルリストツール
' 指定フォルダ内のファイル一覧をハイパーリンク付きでFileListシートに出力する
'=============================================================================

Sub GenerateFileList()
    Dim wsSettings  As Worksheet
    Dim wsResult    As Worksheet
    Dim folderPath  As String
    Dim inclSub     As Boolean
    Dim filterExt   As String

    ' --- エラー発生時の飛び先を指定 ---
    On Error GoTo ErrorHandler

    Set wsSettings = ThisWorkbook.Sheets("Settings")
    Set wsResult = ThisWorkbook.Sheets("FileList")

    '指定フォルダパス
    folderPath = Trim(wsSettings.Range("C4").Value)
    
    'サブフォルダも検索するか否か
    ' Z6セルが 1（はい）なら True、それ以外（いいえ）なら False
    inclSub = (wsSettings.Range("H6").Value = 1)
    
    'ファイル拡張子で絞る（通常はAll）
    filterExt = LCase(Trim(wsSettings.Range("C8").Value))

    If folderPath = "" Then
        MsgBox "フォルダを指定してください.", vbExclamation, "Input Error"
        Exit Sub
    End If

    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' FSOに投げる際、フォルダが存在しない場合、エラー
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(folderPath) Then
        MsgBox "フォルダが存在しません:" & vbCrLf & folderPath, vbCritical, "Folder Error"
        Exit Sub
    End If

    ' FileListシートのセルを全クリア・ハイパーリンク全解除
    wsResult.Cells.Clear
    wsResult.Hyperlinks.Delete
    
    'FileListシートのヘッダ設定
    Call SetupHeaders(wsResult)

    Dim rowIdx As Long
    rowIdx = 3

    'アプリの処理が終わるまで、Excelを書き換えない設定
    Application.ScreenUpdating = False
    
    'FileListシートの詳細部分設定
    Call EnumerateFiles(wsResult, folderPath, rowIdx, inclSub, filterExt)
    
    'FileListシートのフォーマット設定
    Call FormatSheet(wsResult, rowIdx - 1)
    
    wsResult.Activate
    MsgBox "処理終了 合計ファイル数: " & (rowIdx - 3), vbInformation, "ファイルリストツール"
    
CleanUp:
    ' --- 後処理（エラーが起きても必ずここを通る） ---
    Application.ScreenUpdating = True ' 画面更新を必ず再開
    Exit Sub

ErrorHandler:
    ' --- エラー内容の通知 ---
    MsgBox "予期せぬエラーが発生しました." & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "Execution Error"
    Resume CleanUp ' 後処理へ誘導
    
End Sub

'=============================================================================
' FileListシートのヘッダ設定
'=============================================================================
Private Sub SetupHeaders(ws As Worksheet)
    
    ' 1行目の設定
    With ws.Range("A1:G1")
        .Merge
        .Value = "File List"
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Rows(1).RowHeight = 30

    ' 2行目の設定
    Dim headers As Variant
    headers = Array("No.", "File Name", "Extension", "Folder Path", "Size (KB)", "Modified", "Link")
    Dim i As Integer
    For i = 0 To 6
        With ws.Cells(2, i + 1)
            .Value = headers(i)
            .Font.Name = "Arial"
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(68, 114, 196)
            .HorizontalAlignment = xlCenter
        End With
    Next i
    ws.Rows(2).RowHeight = 22

    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 50
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 20
    ws.Columns("G").ColumnWidth = 10

End Sub

'=============================================================================
' FileListシートの詳細部分設定
'=============================================================================
Private Sub EnumerateFiles(ws As Worksheet, currentPath As String, _
                            ByRef rowIdx As Long, inclSub As Boolean, filterExt As String)
    Dim fso    As Object
    Dim folder As Object
    Dim file   As Object
    Dim subF   As Object

    'FSO呼び出し
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(currentPath)

    'ファイルごとに繰り返し処理
    For Each file In folder.Files
        Dim ext As String
        
        '拡張子
        ext = LCase(fso.GetExtensionName(file.Name))

        ' Extension filter
        If filterExt <> "" And filterExt <> "all" Then
            Dim extList As Variant
            extList = Split(filterExt, ",")
            Dim matched As Boolean
            matched = False
            Dim e As Variant
            For Each e In extList
                If Trim(e) = ext Then matched = True
            Next e
            If Not matched Then GoTo NextFile
        End If
    
        ' 通番
        ws.Cells(rowIdx, 1).Value = rowIdx - 2
        'ファイル名
        ws.Cells(rowIdx, 2).Value = file.Name
        '拡張子
        ws.Cells(rowIdx, 3).Value = "." & ext
        '対象フォルダ
        ws.Cells(rowIdx, 4).Value = file.ParentFolder.Path
        'ファイルサイズ
        ws.Cells(rowIdx, 5).Value = Format(file.Size / 1024, "0.00")
        '最終更新日
        ws.Cells(rowIdx, 6).Value = file.DateLastModified
        ws.Cells(rowIdx, 6).NumberFormat = "yyyy/mm/dd hh:mm"

        'ハイパーリンク
        ws.Hyperlinks.Add _
            Anchor:=ws.Cells(rowIdx, 7), _
            Address:=file.Path, _
            TextToDisplay:="Open"
        ws.Cells(rowIdx, 7).HorizontalAlignment = xlCenter

        If rowIdx Mod 2 = 0 Then
            ws.Range(ws.Cells(rowIdx, 1), ws.Cells(rowIdx, 7)).Interior.Color = RGB(218, 227, 243)
        End If

        rowIdx = rowIdx + 1
NextFile:
    Next file

    If inclSub Then
        For Each subF In folder.SubFolders
            Call EnumerateFiles(ws, subF.Path, rowIdx, True, filterExt)
        Next subF
    End If
    
End Sub

'=============================================================================
' FileListシートのフォーマット設定
'=============================================================================
Private Sub FormatSheet(ws As Worksheet, lastRow As Long)
    If lastRow < 3 Then Exit Sub

    '背景スタイル設定
    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 7))
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = RGB(180, 198, 231)
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Color = RGB(180, 198, 231)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
    End With

    '文字の位置設定
    ws.Rows("3:" & lastRow).RowHeight = 18
    ws.Range(ws.Cells(3, 5), ws.Cells(lastRow, 5)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(3, 3), ws.Cells(lastRow, 3)).HorizontalAlignment = xlCenter

    'ウィンドウ枠の固定と、オートフィルタ設定
    ws.Activate
    ws.Range("A3").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A2:G2").AutoFilter

End Sub

'=============================================================================
' FileListシートの全クリア
'=============================================================================
Sub ClearFileList()
    Dim wsResult As Worksheet
    Set wsResult = ThisWorkbook.Sheets("FileList")
    wsResult.Cells.Clear
    wsResult.Hyperlinks.Delete
    MsgBox "FileList sheet cleared.", vbInformation, "FileList Tool"
End Sub

'=============================================================================
' 指定ファイルの参照ボタン押下時処理
'=============================================================================
Sub SelectFolder()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    If fd.Show = -1 Then
        ' 選択されたパスをC4セルに入力
        ThisWorkbook.Sheets("Settings").Range("C4").Value = fd.SelectedItems(1)
    End If
End Sub

'-----------------------------------------------------------------------------
Sub OpenTargetFolder()
    Dim wsSettings As Worksheet
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    Dim folderPath As String
    folderPath = Trim(wsSettings.Range("C4").Value)
    If folderPath = "" Then
        MsgBox "No folder path set in Settings (C4).", vbExclamation
        Exit Sub
    End If
    Shell "explorer.exe """ & folderPath & """", vbNormalFocus
End Sub


