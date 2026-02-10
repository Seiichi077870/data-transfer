Attribute VB_Name = "Module3"
'杉原さん、久保さんのプログラム

Option Explicit
Sub 大中構成チェック()
    Dim folderPath As String
    Dim file1 As String, file2 As String
    Dim filePattern1 As String, filePattern2 As String
    Dim tenDigitNumber1 As String, tenDigitNumber2 As String
    Dim wb1 As Workbook, wb2 As Workbook, wbNew As Workbook
    Dim wsNew1 As Worksheet, wsNew2 As Worksheet
    
    ' フォルダパスの設定
    folderPath = "C:\Temp\大中構成表チェック\"
    
    ' ファイルパターンの設定
    filePattern1 = "*_??????????_*.xlsx"
    filePattern2 = "OBL登録完了_??????????.xlsx"
    
    ' ファイルの存在確認
    file1 = Dir(folderPath & filePattern1)
    file2 = Dir(folderPath & filePattern2)
    
    If file1 = "" Or file2 = "" Then
        MsgBox "必要なファイルが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 10桁の数字部分を抽出
    tenDigitNumber1 = Mid(file1, InStrRev(file1, "_", InStrRev(file1, "_") - 1) + 1, 10)
    tenDigitNumber2 = Mid(file2, InStr(file2, "_") + 1, 10)
    
    ' ファイル名の一致確認
    If tenDigitNumber1 <> tenDigitNumber2 Then
        MsgBox "ファイル名の数字部分が一致しません。", vbExclamation
        Exit Sub
    End If
    
    ' 新しいファイルの作成
    Set wbNew = Workbooks.Add
    wbNew.SaveAs folderPath & "大中構成チェック_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    
    ' 既存ファイルを開く
    Set wb1 = Workbooks.Open(folderPath & file1)
    Set wb2 = Workbooks.Open(folderPath & file2)
    
    ' シートをコピー
    wb1.Sheets(1).Copy After:=wbNew.Sheets(wbNew.Sheets.Count)
    Set wsNew1 = wbNew.Sheets(wbNew.Sheets.Count)
    wsNew1.Name = "大中登録Excel"
    
    wb2.Sheets(1).Copy After:=wbNew.Sheets(wbNew.Sheets.Count)
    Set wsNew2 = wbNew.Sheets(wbNew.Sheets.Count)
    wsNew2.Name = "OBL構成完了"
    
    ' 不要なシートを削除
    Application.DisplayAlerts = False
    wbNew.Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
    
    ' Copy元のファイルを閉じる
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=False
    
    ' 『大中登録Excel』シートの加工
    Call 大中登録Excel加工(wsNew1)
    
    ' 比較処理の実行
    Call 比較処理(wbNew)
    
    ' 「大中登録Excel」シートをアクティブにする
    wbNew.Sheets("大中登録Excel").Activate
    
    ' ファイルの移動と整理
    'Call ファイル整理(folderPath, file1, file2, wbNew)
    
    ' 結果ファイルを保存して閉じる
    On Error Resume Next
    wbNew.Save
    wbNew.Close SaveChanges:=True
    On Error GoTo 0
    
    ' 比較に使用した2つのファイルを閉じる
    On Error Resume Next
    Workbooks(file1).Close SaveChanges:=False
    Workbooks(file2).Close SaveChanges:=False
    On Error GoTo 0
    
    MsgBox "処理が完了しました。", vbInformation
End Sub

Sub 大中登録Excel加工(wsNew1 As Worksheet)
    Dim lastRow As Long, lastCol As Long
    Dim itemRow As Long, colPartNo As Long
    Dim foundCell As Range
    Dim i As Long, j As Long
    
    With wsNew1
        ' セルの統合を解除し、元の値を維持する
        Dim cell As Range
        Dim mergedValue As Variant
        Dim mergedRange As Range
        
        For Each cell In .UsedRange
            If cell.MergeCells Then
                Set mergedRange = cell.MergeArea
                mergedValue = cell.Value
                
                ' セルの統合を解除
                cell.UnMerge
                
                ' 元の値を統合されていたセルに入力
                mergedRange.Value = mergedValue
            End If
        Next cell
        
        ' 項目行と列の特定
        Set foundCell = .Cells.Find(What:="登録部品番号", LookIn:=xlValues, LookAt:=xlWhole)
        If foundCell Is Nothing Then
            MsgBox "登録部品番号の列が見つかりません。", vbExclamation
            Exit Sub
        End If
        colPartNo = foundCell.Column
        itemRow = foundCell.Row
        
        ' 最終行と最終列を特定
        lastRow = .Cells(.Rows.Count, colPartNo).End(xlUp).Row
        lastCol = .Cells(itemRow, .Columns.Count).End(xlToLeft).Column
        
        ' 部品番号<流用元>の列を特定
        Dim colPartNoFlow As Long
        Set foundCell = .Cells.Find(What:="部品番号<流用元>", LookIn:=xlValues, LookAt:=xlWhole)
        If foundCell Is Nothing Then
            MsgBox "部品番号<流用元>の列が見つかりません。", vbExclamation
            Exit Sub
        End If
        colPartNoFlow = foundCell.Column

        ' レベル、員数、PCの列を特定
        Dim colLevel As Long, colQuantity As Long, colPC As Long
        For i = 1 To colPartNoFlow
            Select Case .Cells(itemRow, i).Value
                Case "レベル": colLevel = i
                Case "員数": colQuantity = i
                Case "PC": colPC = i
            End Select
        Next i

        ' 項目行の名称変更と書式設定
        For i = 1 To lastCol + 5
            Dim cellValue As String
            cellValue = .Cells(itemRow, i).Value
            
            ' 部品番号<流用元>の処理
            If i = colPartNoFlow Then
                .Cells(itemRow, i).Value = Replace(cellValue, "<流用元>", vbNewLine & "<流用元>")
            ' 部品番号<流用元>より右の列の処理
            ElseIf i > colPartNoFlow And i <= lastCol Then
                If Right(cellValue, 5) <> "<流用元>" Then
                    .Cells(itemRow, i).Value = cellValue & vbNewLine & "<流用元>"
                End If
            ' 新しく追加された列の処理
            ElseIf i > lastCol Then
                .Cells(itemRow, i).Value = cellValue & vbNewLine & "_Check"
            End If
        Next i

        ' 項目行の書式設定
        With .Rows(itemRow)
            .WrapText = True
            .VerticalAlignment = xlTop
            .HorizontalAlignment = xlCenter
        End With

        ' レベル、員数、PCの列幅を6に設定
        If colLevel > 0 Then .Columns(colLevel).ColumnWidth = 6
        If colQuantity > 0 Then .Columns(colQuantity).ColumnWidth = 6
        If colPC > 0 Then .Columns(colPC).ColumnWidth = 6

        ' その他の列の幅を最小4に設定
        For i = 1 To lastCol + 6
            If i <> colPartNoFlow And i <> colLevel And i <> colQuantity And i <> colPC Then
                If .Columns(i).ColumnWidth < 4 Then
                    .Columns(i).ColumnWidth = 7
                End If
            End If
        Next i
        
        ' データ範囲の全ての行を同じ高さで表示
        .Rows(CStr(itemRow + 1) & ":" & CStr(lastRow)).RowHeight = 15 ' 15は任意の高さ（ポイント単位）
        
        ' オートフィルターを適用
        .Range(.Cells(itemRow, 1), .Cells(lastRow, lastCol + 5)).AutoFilter
        
        ' 新しい列の追加と幅の設定
        For i = 1 To 5
            .Columns(lastCol + i).ColumnWidth = 7
        Next i
        .Cells(itemRow, lastCol + 1).Value = "登録品番" & vbNewLine & "_Check"
        .Cells(itemRow, lastCol + 2).Value = "レベル" & vbNewLine & "_Check"
        .Cells(itemRow, lastCol + 3).Value = "員数" & vbNewLine & "_Check"
        .Cells(itemRow, lastCol + 4).Value = "PC" & vbNewLine & "_Check"
        .Cells(itemRow, lastCol + 5).Value = "加工先" & vbNewLine & "_Check"

        ' 新しく追加された列のデータ範囲に罫線を引く
        With .Range(.Cells(itemRow, lastCol + 1), .Cells(lastRow, lastCol + 5))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With

        ' 表全体の文字設定
        With .Range(.Cells(itemRow, 1), .Cells(lastRow, lastCol + 5))
            .Font.Name = "MS Pゴシック"
            .Font.Size = 10
            .ShrinkToFit = False
        End With
    End With
End Sub

Sub 比較処理(wbNew As Workbook)
    Dim wsNew1 As Worksheet
    Dim wsNew2 As Worksheet
    Dim lastRow As Long, lastRow2 As Long, lastCol As Long
    Dim colPartNo As Long, colRegNo As Long, colLevel As Long, colLevelCheck As Long
    Dim colQuantity As Long, colQuantityCheck As Long, colPC As Long, colPCCheck As Long
    Dim colProcess As Long, colProcessCheck As Long
    Dim colRegNoResult As Long, colLevelResult As Long, colQuantityResult As Long
    Dim colPCResult As Long, colProcessResult As Long
    Dim itemRow As Long
    Dim foundCell As Range
    Dim i As Long, j As Long
    Dim isFound As Boolean
    
    ' ワークシートの設定
    Set wsNew1 = wbNew.Sheets("大中登録Excel")
    Set wsNew2 = wbNew.Sheets("OBL構成完了")
    
    ' 大中登録Excelシートの情報を取得
    With wsNew1
        Set foundCell = .Cells.Find(What:="登録部品番号", LookIn:=xlValues, LookAt:=xlWhole)
        If foundCell Is Nothing Then
            MsgBox "登録部品番号の列が見つかりません。", vbExclamation
            Exit Sub
        End If
        colPartNo = foundCell.Column
        itemRow = foundCell.Row
        lastRow = .Cells(.Rows.Count, colPartNo).End(xlUp).Row
        lastCol = .Cells(itemRow, .Columns.Count).End(xlToLeft).Column

        ' 各列を特定
        colRegNo = .Cells.Find(What:="登録品番" & vbNewLine & "_Check", LookIn:=xlValues, LookAt:=xlWhole).Column
        colLevelCheck = .Cells.Find(What:="レベル" & vbNewLine & "_Check", LookIn:=xlValues, LookAt:=xlWhole).Column
        colLevel = .Cells.Find(What:="レベル", LookIn:=xlValues, LookAt:=xlWhole).Column
        colQuantityCheck = .Cells.Find(What:="員数" & vbNewLine & "_Check", LookIn:=xlValues, LookAt:=xlWhole).Column
        colQuantity = .Cells.Find(What:="員数", LookIn:=xlValues, LookAt:=xlWhole).Column
        colPCCheck = .Cells.Find(What:="PC" & vbNewLine & "_Check", LookIn:=xlValues, LookAt:=xlWhole).Column
        colPC = .Cells.Find(What:="PC", LookIn:=xlValues, LookAt:=xlWhole).Column
        colProcessCheck = .Cells.Find(What:="加工先" & vbNewLine & "_Check", LookIn:=xlValues, LookAt:=xlWhole).Column
        colProcess = .Cells.Find(What:="加工先", LookIn:=xlValues, LookAt:=xlWhole).Column

        ' 新しい列を追加
        lastCol = .Cells(itemRow, .Columns.Count).End(xlToLeft).Column
        colRegNoResult = lastCol + 1
        colLevelResult = lastCol + 2
        colQuantityResult = lastCol + 3
        colPCResult = lastCol + 4
        colProcessResult = lastCol + 5

        ' 新しい列のヘッダーを設定
        .Cells(itemRow, colRegNoResult).Value = "登録品番" & vbNewLine & "_合否"
        .Cells(itemRow, colLevelResult).Value = "レベル" & vbNewLine & "_合否"
        .Cells(itemRow, colQuantityResult).Value = "員数" & vbNewLine & "_合否"
        .Cells(itemRow, colPCResult).Value = "PC" & vbNewLine & "_合否"
        .Cells(itemRow, colProcessResult).Value = "加工先" & vbNewLine & "_合否"

        ' 新しい列のフォーマットを設定
        With .Range(.Cells(itemRow, colRegNoResult), .Cells(lastRow, colProcessResult))
            .Font.Name = "MS Pゴシック"
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .WrapText = True
            .ColumnWidth = 8
        End With

        ' レベル_Checkと加工先_Check列を左寄せに設定
        .Range(.Cells(itemRow + 1, colLevelCheck), .Cells(lastRow, colLevelCheck)).HorizontalAlignment = xlLeft
        .Range(.Cells(itemRow + 1, colProcessCheck), .Cells(lastRow, colProcessCheck)).HorizontalAlignment = xlLeft

        ' OBL構成完了シートの最終行を取得
        lastRow2 = wsNew2.Cells(wsNew2.Rows.Count, "E").End(xlUp).Row

        ' 処理開始メッセージ
        MsgBox "比較処理を開始します。大中登録Excel: " & lastRow & "行, OBL構成完了: " & lastRow2 & "行", vbInformation


        ' 大中登録Excelシートの各部品番号に対して処理
        For i = itemRow + 1 To lastRow
            Dim partNo As String
            partNo = .Cells(i, colPartNo).Value
            
            ' OBL構成完了シートで該当の部品番号を検索
            isFound = False
            For j = 2 To lastRow2 ' ヘッダー行をスキップするため2から開始
                If wsNew2.Cells(j, 5).Value = partNo Then ' E列は5列目
                    isFound = True
                    ' 見つかった行番号を大中登録ExcelシートのP列に記入
                    .Cells(i, colRegNo).Value = j
                    ' OBL構成完了シートの値をCheck列に入力
                    .Cells(i, colLevelCheck).Value = wsNew2.Cells(j, 7).Value  ' G列：レベル
                    .Cells(i, colQuantityCheck).Value = wsNew2.Cells(j, 9).Value  ' I列：員数
                    .Cells(i, colPCCheck).Value = wsNew2.Cells(j, 10).Value  ' J列：PC
                    .Cells(i, colProcessCheck).Value = wsNew2.Cells(j, 11).Value  ' K列：加工先
                    ' OBL構成完了シートの該当行を黄色に塗る
                    wsNew2.Range(wsNew2.Cells(j, 1), wsNew2.Cells(j, wsNew2.Columns.Count)).Interior.Color = RGB(255, 255, 0) ' 黄色
                    Exit For
                End If
            Next j
            
            ' 見つからなかった場合、×を入れる
            If Not isFound Then
                .Cells(i, colRegNo).Value = "×"
                .Cells(i, colLevelCheck).Value = "×"
                .Cells(i, colQuantityCheck).Value = "×"
                .Cells(i, colPCCheck).Value = "×"
                .Cells(i, colProcessCheck).Value = "×"
            End If

            ' 合否判定
            .Cells(i, colRegNoResult).Value = IIf(IsNumeric(.Cells(i, colRegNo).Value), "○", "×")
            .Cells(i, colLevelResult).Value = IIf(Trim(.Cells(i, colLevel).Value) = Trim(.Cells(i, colLevelCheck).Value), "○", "×")

            ' 員数の比較
            On Error Resume Next
            .Cells(i, colQuantityResult).Value = IIf(CDbl(.Cells(i, colQuantity).Value) = CDbl(.Cells(i, colQuantityCheck).Value), "○", "×")
            If Err.Number <> 0 Then
                .Cells(i, colQuantityResult).Value = IIf(Trim(.Cells(i, colQuantity).Value) = Trim(.Cells(i, colQuantityCheck).Value), "○", "×")
            End If
            On Error GoTo 0

            ' PCの比較
            On Error Resume Next
            .Cells(i, colPCResult).Value = IIf(CDbl(.Cells(i, colPC).Value) = CDbl(.Cells(i, colPCCheck).Value), "○", "×")
            If Err.Number <> 0 Then
                .Cells(i, colPCResult).Value = IIf(Trim(.Cells(i, colPC).Value) = Trim(.Cells(i, colPCCheck).Value), "○", "×")
            End If
            On Error GoTo 0

            ' 加工先の比較
            On Error Resume Next
            .Cells(i, colProcessResult).Value = IIf(CDbl(Replace(.Cells(i, colProcess).Value, " ", "")) = CDbl(Replace(.Cells(i, colProcessCheck).Value, " ", "")), "○", "×")
            If Err.Number <> 0 Then
                .Cells(i, colProcessResult).Value = IIf(Trim(Replace(.Cells(i, colProcess).Value, " ", "")) = Trim(Replace(.Cells(i, colProcessCheck).Value, " ", "")), "○", "×")
            End If
            On Error GoTo 0

            ' ×のセルを赤く塗る
            For j = colRegNoResult To colProcessResult
                If .Cells(i, j).Value = "×" Then
                    .Cells(i, j).Interior.Color = RGB(255, 0, 0)
                End If
            Next j
        Next i

        ' OBL構成完了シートの部品番号が大中登録Excelシートに存在しない場合の処理
        For j = 2 To lastRow2
            Dim oblPartNo As String
            oblPartNo = wsNew2.Cells(j, 5).Value
            
            isFound = False
            For i = itemRow + 1 To lastRow
                If .Cells(i, colPartNo).Value = oblPartNo Then
                    isFound = True
                    Exit For
                End If
            Next i
            
            If Not isFound Then
                wsNew2.Range(wsNew2.Cells(j, 1), wsNew2.Cells(j, wsNew2.Columns.Count)).Interior.Color = RGB(169, 169, 169) ' 濃いグレー
            End If
        Next j
        
        ' グループ化と非表示の処理
        Dim groupStart As Long, groupEnd As Long
        Dim checkGroupStart As Long, checkGroupEnd As Long
        Dim separatorColumn As Long
        groupStart = 0
        groupEnd = 0
        checkGroupStart = 0
        checkGroupEnd = 0
        
        ' 既存のグループをすべて解除
        On Error Resume Next
        .Outline.ShowLevels ColumnLevels:=1
        On Error GoTo 0
        
        ' すべての列の非表示を解除
        .Columns.Hidden = False
        
        For i = 1 To lastCol + 5
            Dim cellValue As String
            cellValue = .Cells(itemRow, i).Value
            
            If InStr(cellValue, "流用元") > 0 Then
                If groupStart = 0 Then groupStart = i
                groupEnd = i
            ElseIf InStr(cellValue, "_Check") > 0 Then
                If checkGroupStart = 0 Then
                    ' 区切り列を挿入
                    .Columns(i).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                    separatorColumn = i
                    i = i + 1 ' 列が挿入されたので、インデックスを1つ増やす
                    checkGroupStart = i
                End If
                checkGroupEnd = i
            Else
                ' 流用元グループの処理
                If groupStart > 0 And groupEnd > 0 Then
                    On Error Resume Next
                    .Columns(groupStart & ":" & groupEnd).Group
                    If Err.Number = 0 Then
                        .Columns(groupStart & ":" & groupEnd).Hidden = True
                    End If
                    On Error GoTo 0
                    groupStart = 0
                    groupEnd = 0
                End If
                ' _Checkグループの処理
                If checkGroupStart > 0 And checkGroupEnd > 0 Then
                    On Error Resume Next
                    .Columns(checkGroupStart & ":" & checkGroupEnd).Group
                    On Error GoTo 0
                    checkGroupStart = 0
                    checkGroupEnd = 0
                End If
            End If
        Next i
        
        ' 最後のグループ処理
        If groupStart > 0 And groupEnd > 0 Then
            On Error Resume Next
            .Columns(groupStart & ":" & groupEnd).Group
            If Err.Number = 0 Then
                .Columns(groupStart & ":" & groupEnd).Hidden = True
            End If
            On Error GoTo 0
        End If
        If checkGroupStart > 0 And checkGroupEnd > 0 Then
            On Error Resume Next
            .Columns(checkGroupStart & ":" & checkGroupEnd).Group
            On Error GoTo 0
        End If
        
        ' 区切り列の書式設定
        If separatorColumn > 0 Then
            With .Columns(separatorColumn)
                .ColumnWidth = 0.1 ' 非常に狭い幅に設定
                .Interior.Color = RGB(0, 0, 0) ' 黒色に設定
            End With
        End If
        
        ' グループの展開レベルを設定
        .Outline.ShowLevels ColumnLevels:=2
        
        ' シートを表示状態に設定
        .Visible = xlSheetVisible
    End With

    ' 処理完了メッセージ
    MsgBox "比較処理が完了しました。", vbInformation
End Sub

Sub ファイル整理(folderPath As String, file1 As String, file2 As String, wbNew As Workbook)
    Dim tenDigitNumber As String
    Dim newFolderPath As String
    Dim fso As Object
    Dim suffix As Integer
    
    ' FileSystemObjectの作成
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 10桁の数字を抽出
    tenDigitNumber = Mid(file1, InStrRev(file1, "_", InStrRev(file1, "_") - 1) + 1, 10)
    
    ' 新しいフォルダパスの作成
    newFolderPath = folderPath & "確認済フォルダ\" & tenDigitNumber & "_" & Format(Date, "yyyymmdd")
    
    ' 同名のフォルダが存在する場合、末尾に連番を付ける
    suffix = 1
    Do While fso.FolderExists(newFolderPath)
        suffix = suffix + 1
        newFolderPath = folderPath & "確認済フォルダ\" & tenDigitNumber & "_" & Format(Date, "yyyymmdd") & "_" & suffix
    Loop
    
    ' フォルダを作成
    fso.CreateFolder newFolderPath
    
    ' ファイルの移動
    fso.MoveFile folderPath & file1, newFolderPath & "\" & file1
    fso.MoveFile folderPath & file2, newFolderPath & "\" & file2
    
    ' 「大中登録Excel」シートをアクティブにしてから保存
    wbNew.Sheets("大中登録Excel").Activate
    wbNew.SaveAs newFolderPath & "\" & wbNew.Name
    ' ここでwbNewを閉じない
    
    ' FileSystemObjectの解放
    Set fso = Nothing
End Sub

Function IsAlphaNumeric(char As String) As Boolean
    ' 半角英数字のみを許可
    IsAlphaNumeric = (Asc(char) >= 48 And Asc(char) <= 57) Or _
                     (Asc(char) >= 65 And Asc(char) <= 90) Or _
                     (Asc(char) >= 97 And Asc(char) <= 122)
End Function



