Option Explicit
Private Enum eRow
    CheckStart = 4          '開始行
    FontName = 11           'フォント名
    FontSize = 12           'フォントサイズ
    Start = 18              '開始行
End Enum
Private Enum eCol
    InputData = 3           '入力列
    Title = 4               'タイトル
    Sheet_Count = 5         'シート数
    Sheet_Name = 6          'シート名
    New_Type = 7            '新形式
    Print_Area = 8          '印刷範囲
    Print_Wakugai = 9       '印刷範囲外ゴミ
    Shape_Picture = 10      'シェイプ（画像）
    Shape_Sonota = 11       'シェイプ（画像以外）
    Name_Name = 12          '名前の管理（名前）
    Name_SansyouArea = 13   '名前の管理（参照範囲）
    Name_Area = 14          '名前の管理（範囲）
End Enum
Private Enum eCheck
    Format = 1              '表示形式
    Dataumu = 2             'データ有無
    Horizontal = 3          '横位置
    Vertical = 4            '縦位置
    Wrap = 5                '折り返し縮小
    Merge = 6               'マージ
    Font = 7                'フォント
End Enum
Private Enum eColor
    none = 0
    red = 3
    green = 4
    blue = 33
    pink = 26
    yellow = 27
End Enum
Private Const msCHECK_ON As String = "チェックする"
Private msFont_Name As String
Private mnFont_Size As Integer

Private Sub cmdExcelGet_Click()
Dim sFoldaPath As String
Dim nRow As Long
Dim nCol As Long
On Error GoTo Err_
    'フォルダ選択ダイアログ表示
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = False Then
            Exit Sub
        End If
        sFoldaPath = .SelectedItems(1)
    End With
    Application.ScreenUpdating = False
    'セル初期化
    nRow = Range("C65536").End(xlUp).Row
    If nRow < eRow.Start Then
        nRow = eRow.Start
    End If
    nCol = UsedRange.Columns(UsedRange.Columns.Count).Column
    Range(Cells(eRow.Start, eCol.InputData), Cells(nRow, nCol)).ClearContents
    'エクセルファイル名列挙
    Call subGetExcelFile(sFoldaPath)
    Application.ScreenUpdating = True
    Exit Sub
Err_:
    MsgBox ("エラーが発生しました。" & vbNewLine & "エラー番号: " & Err.Number & vbNewLine & "エラー内容: " & Err.Description & vbNewLine)
    Application.ScreenUpdating = True
End Sub

Private Sub subGetExcelFile(ByVal sFoldaPath As String)
Dim nRow As Long
Dim sFileName As String
Dim oFolder As Object
On Error GoTo Err_
    If Right(sFoldaPath, 1) <> "\" Then
        sFoldaPath = sFoldaPath & "\"
    End If
    nRow = Range("C65536").End(xlUp).Row + 1
    If nRow < eRow.Start Then
        nRow = eRow.Start
    End If
    'フォルダ内のエクセルファイルを検索
    sFileName = Dir(sFoldaPath & "*.xlsx")
    Do While sFileName <> ""
        Cells(nRow, eCol.InputData) = sFoldaPath & sFileName
        sFileName = Dir()
        nRow = nRow + 1
    Loop
    'フォルダ内のサブフォルダを検索
    With CreateObject("Scripting.FileSystemObject")
        For Each oFolder In .GetFolder(sFoldaPath).SubFolders
            Call subGetExcelFile(oFolder.Path)
        Next
    End With
    Exit Sub
Err_:
    MsgBox ("エラーが発生しました。" & vbNewLine & "エラー番号: " & Err.Number & vbNewLine & "エラー内容: " & Err.Description & vbNewLine)
End Sub

Private Sub cmdJikkou_Click()
Dim nRow As Long
Dim nCol As Long
Dim nMaxRow As Long
Dim nUsedMaxRow As Long
Dim nUsedMaxCol As Long
Dim nCheck As Integer
Dim sFilePath As String
Dim nNo As Integer
Dim sSheets() As String
Dim i, j As Long
On Error GoTo Err_
    'エラーチェック
    If Cells(eRow.FontName, eCol.InputData) = "" Then
        MsgBox ("フォント名を指定してください。")
        Exit Sub
    End If
    If Trim(Cells(eRow.FontSize, eCol.InputData)) = "" Or IsNumeric(Cells(eRow.FontSize, eCol.InputData)) = False Then
        MsgBox ("フォントサイズが不正です。")
        Exit Sub
    End If
    For i = eRow.Start To Range("C65536").End(xlUp).Row
        If Dir(Cells(i, eCol.InputData)) = "" Then
            MsgBox ("ファイルがありません。" & vbCr & sFilePath)
            Exit Sub
        End If
    Next
    If MsgBox("実行しますか？", vbOKCancel) = vbCancel Then
        Exit Sub
    End If
    msFont_Name = Cells(eRow.FontName, eCol.InputData)  'フォント名
    mnFont_Size = Cells(eRow.FontSize, eCol.InputData)  'フォントサイズ
    Application.ScreenUpdating = False
    
    '不要シート削除
    Call subSheetDelete
    
    'チェックファイル数でループ
    nMaxRow = Range("C65536").End(xlUp).Row
    For i = eRow.Start To nMaxRow
        Call subWriteBookInf(i)   'ブック情報を取得し、一覧に記載する
        nNo = nNo + 1
        Application.StatusBar = "実行中... " & CStr(i - eRow.Start + 1) & "/" & CStr(nMaxRow - eRow.Start + 1)
        sFilePath = Cells(i, eCol.InputData)
        sSheets = Split(Cells(i, eCol.Sheet_Name), vbNewLine)
        'シート数でループ
        For j = 0 To UBound(sSheets)
            'チェック項目数でループ
            For nCheck = 1 To 7
                If Cells(eRow.CheckStart + nCheck - 1, eCol.InputData) = msCHECK_ON Then 'チェックする場合
                    Call subSheetCopy(sFilePath, sSheets(j))       'シートコピー
                    Call subChangeSheetName(nCheck, nNo, j)        'シート名変更＆Ａ１コメント挿入
                    With ActiveSheet.UsedRange
                        nUsedMaxRow = .Rows(.Rows.Count).Row           '最大行番号
                        nUsedMaxCol = .Columns(.Columns.Count).Column  '最大列番号
                    End With
                    Call subBeforeCheckCell(nCheck, nUsedMaxRow, nUsedMaxCol)
                    For nRow = 1 To nUsedMaxRow + 10        '枠外に+10行
                        For nCol = 1 To nUsedMaxCol + 10    '枠外に+10列
                            Call subCheckCell(nCheck, nRow, nCol, nUsedMaxRow, nUsedMaxCol) 'セルチェック
                        Next
                    Next
                End If
            Next
        Next
    Next
    Sheets(Me.Name).Select
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub
Err_:
    MsgBox ("エラーが発生しました。" & vbNewLine & "エラー番号: " & Err.Number & vbNewLine & "エラー内容: " & Err.Description & vbNewLine)
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

Private Sub subWriteBookInf(ByVal nRow As Integer)
Dim oSheet As Worksheet
Dim oObj As Object
Dim sFilePath As String
Dim sSheetName As String
Dim sPrintArea As String
Dim sPrintWakugai As String
Dim sShapePicture As String
Dim sShapeSonota As String
Dim sNameName As String
Dim sNameSansyouArea
Dim sNameArea As String
Dim bNewType As Boolean
On Error GoTo Err_
    sFilePath = Cells(nRow, eCol.InputData)
    Workbooks.Open sFilePath
    With Workbooks(Dir(sFilePath))
        .CheckCompatibility = False
        Cells(nRow, eCol.Title) = .BuiltinDocumentProperties("Title").Value 'タイトル
        Cells(nRow, eCol.Sheet_Count) = .Sheets.Count       'シート数
        'ブック内のシートでループ
        For Each oSheet In .Worksheets
            If sSheetName <> "" Then
                sSheetName = sSheetName & vbNewLine
                sPrintArea = sPrintArea & vbNewLine
                sPrintWakugai = sPrintWakugai & vbNewLine
            End If
            sSheetName = sSheetName & oSheet.Name
            sPrintArea = sPrintArea & oSheet.PageSetup.PrintArea
            sPrintWakugai = sPrintWakugai & fncCheckPrintWakugai(oSheet)

            '全てのセルよりモデル名を検索
            If oSheet.Cells.Find("@*@") Is Nothing = False Then
                bNewType = True
            End If

            'ブック内のシェイプでループ
            For Each oObj In oSheet.Shapes
                If oObj.Type = msoPicture Then
                    '画像
                    If sShapePicture <> "" Then
                        sShapePicture = sShapePicture & vbNewLine
                    End If
                    sShapePicture = sShapePicture & oObj.Name
                Else
                    '画像以外
                    If sShapeSonota <> "" Then
                        sShapeSonota = sShapeSonota & vbNewLine
                    End If
                    sShapeSonota = sShapeSonota & oObj.Name
                End If
            Next
        Next
        Cells(nRow, eCol.Sheet_Name) = sSheetName       'シート名
        If bNewType = True Then
            Cells(nRow, eCol.New_Type) = "○"           '新形式
        End If
        Cells(nRow, eCol.Print_Area) = sPrintArea       '印刷範囲
        Cells(nRow, eCol.Print_Wakugai) = sPrintWakugai '印刷範囲外ゴミ
        Cells(nRow, eCol.Shape_Picture) = sShapePicture 'シェイプ名（画像）
        Cells(nRow, eCol.Shape_Sonota) = sShapeSonota   'シェイプ名（画像以外）
        
        'ブック内の名前の定義でループ
        For Each oObj In .Names
            If sNameName <> "" Then
                sNameName = sNameName & vbNewLine
                sNameSansyouArea = sNameSansyouArea & vbNewLine
                sNameArea = sNameArea & vbNewLine
            End If
            sNameName = sNameName & oObj.Name
            sNameSansyouArea = sNameSansyouArea & oObj.RefersTo
            sNameArea = sNameArea & oObj.Parent.Name
        Next
        Cells(nRow, eCol.Name_Name) = "'" & sNameName               '名前管理（名前）
        Cells(nRow, eCol.Name_SansyouArea) = "'" & sNameSansyouArea '名前管理（参照範囲）
        Cells(nRow, eCol.Name_Area) = "'" & sNameArea               '名前管理（範囲）
        .Close False
    End With
    Exit Sub
Err_:
    MsgBox ("エラーが発生しました。" & vbNewLine & "エラー番号: " & Err.Number & vbNewLine & "エラー内容: " & Err.Description & vbNewLine)
End Sub

Private Function fncCheckPrintWakugai(ByVal oSheet As Worksheet) As String
Dim nMaxRow As Long
Dim nMaxCol As Long
Dim sPrintArea As String
Dim nPrintMaxRow As Long
Dim nPrintMaxCol As Long
Dim sArr1() As String
Dim sArr2() As String
Dim sRange1 As String
Dim sRange2 As String
On Error GoTo Err_
    fncCheckPrintWakugai = "不明"
    sPrintArea = oSheet.PageSetup.PrintArea
    If sPrintArea = "" Then
        Exit Function
    End If
    sArr1 = Split(sPrintArea, ":")
    If UBound(sArr1) = 0 Then
        sRange1 = sArr1(0)
    ElseIf UBound(sArr1) = 1 Then
        sRange1 = sArr1(1)
    Else
        Exit Function
    End If
    sArr2 = Split(sRange1, "$")
    If UBound(sArr2) <> 2 Then
        Exit Function
    End If
    sRange2 = Replace(sRange1, "$", "")
    nPrintMaxRow = oSheet.Range(sRange2).Row
    nPrintMaxCol = oSheet.Range(sRange2).Column
    With oSheet.UsedRange
        nMaxRow = .Rows(.Rows.Count).Row           '最大行番号
        nMaxCol = .Columns(.Columns.Count).Column  '最大列番号
    End With
    If nPrintMaxRow < nMaxRow Or nPrintMaxCol < nMaxCol Then
        fncCheckPrintWakugai = "有り"
    Else
        fncCheckPrintWakugai = "無し"
    End If
    Exit Function
Err_:
    MsgBox ("エラーが発生しました。" & vbNewLine & "エラー番号: " & Err.Number & vbNewLine & "エラー内容: " & Err.Description & vbNewLine)
End Function

Private Sub subBeforeCheckCell(ByVal nCheck As Integer, ByVal nMaxRow As Long, ByVal nMaxCol As Long)
On Error GoTo Err_
    Select Case nCheck
    Case eCheck.Dataumu 'データ有無
        '最大の行と列に背景色黄色を設定
        ActiveSheet.Rows(nMaxRow).Interior.ColorIndex = eColor.yellow
        ActiveSheet.Columns(nMaxCol).Interior.ColorIndex = eColor.yellow
    End Select
    Exit Sub
Err_:
    MsgBox ("エラーが発生しました。" & vbNewLine & "エラー番号: " & Err.Number & vbNewLine & "エラー内容: " & Err.Description & vbNewLine)
End Sub

Private Sub subChangeSheetName(ByVal nCheck As Integer, ByVal nNo As Integer, ByVal nSheetCount As Integer)
Dim sSheetName As String
Dim sComment As String
On Error GoTo Err_
    With ActiveSheet
        Select Case nCheck
        Case eCheck.Format
            sSheetName = CStr(nNo) & ".表示形式"
            sComment = "青：文字列" & vbNewLine & "緑：数値" & vbNewLine & "無：標準" & vbNewLine & "赤：枠外の標準以外" & vbNewLine & "黄：ユーザ定義"
        Case eCheck.Dataumu
            sSheetName = CStr(nNo) & ".セルデータ有無"
            sComment = "赤：値あり" & vbNewLine & "黄：最大行列番号"
        Case eCheck.Horizontal
            sSheetName = CStr(nNo) & ".横位置"
            sComment = "青：左" & vbNewLine & "黄：中央" & vbNewLine & "緑：右" & vbNewLine & "赤：その他" & vbNewLine & "無：指定なし"
        Case eCheck.Vertical
            sSheetName = CStr(nNo) & ".縦位置"
            sComment = "青：上" & vbNewLine & "黄：中央" & vbNewLine & "緑：下" & vbNewLine & "赤：その他"
        Case eCheck.Wrap
            sSheetName = CStr(nNo) & ".折縮小"
            sComment = "青：折り返し" & vbNewLine & "改行位置で行が分かれて表示される。" & vbNewLine & "文字数が多いと文字切れが発生する。" & vbNewLine & vbNewLine
            sComment = sComment & "黄：縮小" & vbNewLine & "改行位置で行が分かれず、セル内に収まるよう表示される。" & vbNewLine & "文字切れは発生しない。"
        Case eCheck.Merge
            sSheetName = CStr(nNo) & ".マージ"
            sComment = "青：マージ"
        Case eCheck.Font
            sSheetName = CStr(nNo) & ".フォント"
            sComment = "赤：" & msFont_Name & "以外" & vbNewLine & "青：" & CStr(mnFont_Size) & "pt" & vbNewLine & "網掛：太字・黒以外"
        End Select
        If nSheetCount <> 0 Then
            sSheetName = sSheetName & "(" & CStr(nSheetCount + 1) & ")"
        End If
        .Name = sSheetName
        .Cells(1, 1).ClearComments
        .Cells(1, 1).AddComment (sComment)                  'Ａ１コメント
        .Cells(1, 1).Comment.Shape.Width = 300              'コメント幅
        .Cells(1, 1).Comment.Shape.Height = 100             'コメント高さ
    End With
    Exit Sub
Err_:
    MsgBox ("エラーが発生しました。" & vbNewLine & "エラー番号: " & Err.Number & vbNewLine & "エラー内容: " & Err.Description & vbNewLine)
End Sub

Private Sub subCheckCell(ByVal nCheck As Integer, ByVal nRow As Long, ByVal nCol As Long, ByVal nMaxRow As Long, ByVal nMaxCol As Long)
Dim nColor As eColor
Dim nPattern As Integer
On Error GoTo Err_
    nColor = eColor.none
    nPattern = xlSolid
    With ActiveSheet.Cells(nRow, nCol)
        Select Case nCheck
        Case eCheck.Format
            '表示形式
            If nCol <= nMaxCol And nRow <= nMaxRow Then
                '枠内
                Select Case Trim(.NumberFormatLocal)
                Case "G/標準"       '標準
                    nColor = eColor.none
                Case "@"            '文字列
                    nColor = eColor.blue
                Case "0_", "#,##0_", "0.0_" '数値（単純パターン）
                    nColor = eColor.green
                Case Else
                    nColor = eColor.yellow
                End Select
            Else
                '枠外
                If Trim(.NumberFormatLocal) = "G/標準" Then
                    nColor = eColor.none
                Else
                    nColor = eColor.red
                End If
            End If
        Case eCheck.Dataumu
            'データ有無
            If .HasFormula = True Then  '数式の場合
                nColor = eColor.red
            Else
                If .Value <> "" Then
                    nColor = eColor.red
                End If
            End If
        Case eCheck.Horizontal
            '横位置
            Select Case .HorizontalAlignment
            Case xlLeft         '左詰め
                nColor = eColor.blue
            Case xlJustify      '両端揃え
                nColor = eColor.red
            Case xlCenter       '中央揃え
                nColor = eColor.yellow
            Case xlDistributed  '均等割り付け
                nColor = eColor.red
            Case xlFill         '繰り返し
                nColor = eColor.red
            Case xlRight        '右詰め
                nColor = eColor.green
            Case xlCenterAcrossSelection    '選択範囲内で中央
                nColor = eColor.red
            Case Else   '標準
            End Select
        Case eCheck.Vertical
            '縦位置
            Select Case .VerticalAlignment
            Case xlTop          '上詰め
                nColor = eColor.blue
            Case xlJustify      '両端揃え
                nColor = eColor.red
            Case xlCenter       '中央揃え
                nColor = eColor.yellow
            Case xlDistributed  '均等割り付け
                nColor = eColor.blue
            Case xlBottom       '下詰め
                nColor = eColor.green
            End Select
        Case eCheck.Wrap
            '折り返し・縮小
            If .WrapText = True Then        '折り返し
                nColor = eColor.blue
            Else
                If .ShrinkToFit = True Then '縮小
                    nColor = eColor.yellow
                End If
            End If
        Case eCheck.Merge
            'マージ
            If .MergeCells = True Then
                nColor = eColor.blue
            End If
        Case eCheck.Font
            'フォント
            If .Font.Size = mnFont_Size Then   'サイズ
                nColor = eColor.blue
            End If
            If .Font.Name <> msFont_Name Then
                nColor = eColor.red
            End If
            If .Font.Bold = True Then           '太字
                nPattern = xlGrid
            End If
            If .Font.Color <> 0 Then            'カラー（黒以外）
                nPattern = xlGrid
            End If
        End Select
        '色
        If nColor <> eColor.none Then
            .Interior.ColorIndex = nColor
        End If
        '網掛け
        If nPattern <> xlSolid Then
            .Interior.Pattern = nPattern
        End If
    End With
    Exit Sub
Err_:
    MsgBox ("エラーが発生しました。" & vbNewLine & "エラー番号: " & Err.Number & vbNewLine & "エラー内容: " & Err.Description & vbNewLine)
End Sub

Private Sub subSheetCopy(ByVal sFilePath As String, ByVal sSheetName As String)
Dim sFileName As String
On Error GoTo Err_
    sFileName = Dir(sFilePath)
    Workbooks.Open sFilePath
    With Workbooks(sFileName)
        .CheckCompatibility = False
        .Sheets(sSheetName).Copy After:=Workbooks(ThisWorkbook.Name).Sheets(Workbooks(ThisWorkbook.Name).Sheets.Count)
        .Close False
    End With
    Exit Sub
Err_:
    MsgBox ("エラーが発生しました。" & vbNewLine & "エラー番号: " & Err.Number & vbNewLine & "エラー内容: " & Err.Description & vbNewLine)
End Sub

Private Sub subSheetDelete()
Dim i As Integer
On Error GoTo Err_
    Application.DisplayAlerts = False
    For i = Sheets.Count To 1 Step -1
        If Sheets(i).Name <> Me.Name Then   '自身以外をシート削除
            Sheets(i).Delete
        End If
    Next
    Application.DisplayAlerts = True
    Exit Sub
Err_:
    MsgBox ("エラーが発生しました。" & vbNewLine & "エラー番号: " & Err.Number & vbNewLine & "エラー内容: " & Err.Description & vbNewLine)
End Sub
