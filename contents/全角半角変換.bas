Attribute VB_Name = "全角半角変換"
Sub ZenkakuHankakuConverter()

    Dim ws As Worksheet
    Dim rng As Range
    Dim frm As Object
    
    ' エラーハンドリング
    On Error GoTo ErrorHandler
    
    ' アクティブなワークシート取得
    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "アクティブなワークシートがありません。", vbExclamation
        Exit Sub
    End If
    
    ' シートが保護されているかチェック
    If ws.ProtectContents Then
        MsgBox "シート「" & ws.Name & "」が保護されています。保護を解除してから実行してください。", vbExclamation
        Exit Sub
    End If
    
    ' 選択範囲の確認と設定
    Set rng = selection
    If rng Is Nothing Then
        MsgBox "範囲が選択されていません。", vbExclamation
        Exit Sub
    End If
    
    ' 単一の空セルが選択されている場合
    If rng.count = 1 And IsEmpty(rng) Then
        If MsgBox("選択されているセルが空です。" & vbCrLf & _
                 "使用されている全範囲を対象にしますか？", _
                 vbYesNo + vbQuestion) = vbYes Then
            Set rng = ws.UsedRange
        Else
            Exit Sub
        End If
    End If
    
    ' フォームを表示して設定取得
    Dim useForm As Boolean
    useForm = True
    
    ' フォームの作成を試行
    On Error Resume Next
    Set frm = VBA.UserForms.Add("frmZenkakuHankaku")
    If Err.Number <> 0 Or frm Is Nothing Then
        useForm = False
    End If
    On Error GoTo 0
    
    If useForm Then
        ' フォームが利用可能な場合
        On Error Resume Next
        frm.Show
        If Err.Number <> 0 Then
            ' フォーム表示でエラーが発生した場合
            useForm = False
            Unload frm
            Set frm = Nothing
        End If
        On Error GoTo 0
        
        If useForm Then
            If Not frm.ProcessExecuted Then
                ' キャンセルまたは×で閉じられた場合は終了
                Unload frm
                Set frm = Nothing
                Exit Sub
            End If
            
            ' 設定を取得
            Dim convDirection As Integer
            Dim convAlphaNumeric As Boolean
            Dim convSymbols As Boolean
            Dim convKatakana As Boolean
            Dim convSpaces As Boolean
            Dim IncludeFormulas As Boolean
            
            convDirection = frm.ConversionDirection
            convAlphaNumeric = frm.ConvertAlphaNumeric
            convSymbols = frm.ConvertSymbols
            convKatakana = frm.ConvertKatakana
            convSpaces = frm.ConvertSpaces
            IncludeFormulas = frm.IncludeFormulas
            
            ' フォームを閉じる
            Unload frm
            Set frm = Nothing
            
            ' 変換実行
            Call ExecuteConversion(rng, convDirection, convAlphaNumeric, convSymbols, convKatakana, convSpaces, IncludeFormulas)
        End If
    End If
    
    If Not useForm Then
        ' フォームが利用できない場合は簡易版を使用
        If Not GetSettingsSimple(rng) Then
            Exit Sub
        End If
    End If
    
ErrorHandler:
    ' エラー番号0の場合は何もしない
    If Err.Number = 0 Then
        Exit Sub
    End If
    
    ' フォームのクリーンアップ
    On Error Resume Next
    If Not frm Is Nothing Then
        Unload frm
        Set frm = Nothing
    End If
    On Error GoTo 0
    
    ' エラー番号と内容が有効な場合のみメッセージを表示
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました。" & vbCrLf & _
               "エラー番号: " & Err.Number & vbCrLf & _
               "エラー内容: " & Err.description, vbCritical
    End If
End Sub

Private Function GetSettingsSimple(rng As Range) As Boolean
    Dim direction As Integer
    Dim alphaNumeric As Boolean, symbols As Boolean, katakana As Boolean, spaces As Boolean
    Dim IncludeFormulas As Boolean
    
    On Error Resume Next
    
    ' 変換方向選択
    If MsgBox("変換方向を選択してください" & vbCrLf & vbCrLf & _
             "「はい」: 全角 → 半角" & vbCrLf & _
             "「いいえ」: 半角 → 全角", _
             vbYesNo + vbQuestion, "変換方向選択") = vbYes Then
        direction = 1
    Else
        direction = 2
    End If
    
    ' 変換対象選択
    alphaNumeric = (MsgBox("英数字を変換しますか？", vbYesNo + vbQuestion) = vbYes)
    symbols = (MsgBox("記号を変換しますか？", vbYesNo + vbQuestion) = vbYes)
    katakana = (MsgBox("カタカナを変換しますか？", vbYesNo + vbQuestion) = vbYes)
    spaces = (MsgBox("スペースを変換しますか？", vbYesNo + vbQuestion) = vbYes)
    IncludeFormulas = (MsgBox("数式セルも処理しますか？", vbYesNo + vbQuestion) = vbYes)
    
    ' 最低1つの変換対象が選択されているかチェック
    If Not (alphaNumeric Or symbols Or katakana Or spaces) Then
        MsgBox "変換対象が選択されていません。", vbExclamation
        GetSettingsSimple = False
        On Error GoTo 0
        Exit Function
    End If
    
    On Error GoTo 0
    
    ' 変換実行
    Call ExecuteConversion(rng, direction, alphaNumeric, symbols, katakana, spaces, IncludeFormulas)
    GetSettingsSimple = True
End Function

Private Sub ExecuteConversion(rng As Range, direction As Integer, _
                            alphaNumeric As Boolean, symbols As Boolean, _
                            katakana As Boolean, spaces As Boolean, _
                            IncludeFormulas As Boolean)
    
    Dim cell As Range
    Dim originalValue As Variant
    Dim convertedValue As String
    Dim processedCount As Long
    Dim changedCount As Long
    Dim startTime As Double
    
    startTime = Timer
    
    On Error GoTo ConversionError
    
    ' 画面更新を停止して処理速度向上
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' メイン処理
    For Each cell In rng
        If ShouldProcessCell(cell, IncludeFormulas) Then
            originalValue = cell.value
            Dim originalFormat As String
            originalFormat = cell.NumberFormat
            
            ' 数値の場合は文字列として処理
            If IsNumeric(originalValue) And Not IsEmpty(originalValue) Then
                ' セルの表示形式を文字列に変更
                cell.NumberFormat = "@"
                ' 値を文字列として再設定
                cell.value = CStr(originalValue)
                originalValue = CStr(originalValue)
            End If
            
            convertedValue = ConvertText(CStr(originalValue), direction, alphaNumeric, symbols, katakana, spaces)
            
            If CStr(originalValue) <> convertedValue Then
                ' 変換後の値を設定
                cell.NumberFormat = "@"
                cell.value = convertedValue
                
                ' 全角→半角で数値のみの場合は数値形式に戻す
                If direction = 1 And IsNumericOnly(convertedValue) Then
                    ' 数値として認識できる場合は数値形式に戻す
                    On Error Resume Next
                    Dim numValue As Double
                    numValue = CDbl(convertedValue)
                    If Err.Number = 0 Then
                        cell.NumberFormat = "General"
                        cell.value = numValue
                    End If
                    On Error GoTo ConversionError
                End If
                
                changedCount = changedCount + 1
            End If
            
            processedCount = processedCount + 1
        End If
    Next cell
    
    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' 処理結果を報告
    Dim resultMsg As String
    Dim processingTime As Double
    processingTime = Timer - startTime
    
    Dim directionText As String
    directionText = IIf(direction = 1, "全角→半角", "半角→全角")
    
    resultMsg = "全角半角統一処理が完了しました。" & vbCrLf & vbCrLf & _
               "処理結果:" & vbCrLf & _
               "? 変換方向: " & directionText & vbCrLf & _
               "? 処理セル数: " & Format(processedCount, "#,##0") & vbCrLf & _
               "? 変更セル数: " & Format(changedCount, "#,##0") & vbCrLf & _
               "? 処理時間: " & Format(processingTime, "0.00") & "秒"
    
    MsgBox resultMsg, vbInformation, "処理完了"
    
    Exit Sub
    
ConversionError:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "変換処理中にエラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.description, vbCritical
End Sub

Private Function IsNumericOnly(value As String) As Boolean
    On Error Resume Next
    
    ' 空文字列は数値ではない
    If Len(Trim(value)) = 0 Then
        IsNumericOnly = False
        Exit Function
    End If
    
    ' 数値として変換できるかチェック
    Dim testValue As Double
    testValue = CDbl(value)
    
    If Err.Number = 0 Then
        ' さらに文字列が数字、小数点、符号のみで構成されているかチェック
        Dim i As Integer
        Dim char As String
        For i = 1 To Len(value)
            char = Mid(value, i, 1)
            If Not (char >= "0" And char <= "9") And char <> "." And char <> "-" And char <> "+" Then
                IsNumericOnly = False
                Exit Function
            End If
        Next i
        IsNumericOnly = True
    Else
        IsNumericOnly = False
    End If
    
    On Error GoTo 0
End Function

Private Function ConvertText(inputText As String, direction As Integer, _
                           alphaNumeric As Boolean, symbols As Boolean, _
                           katakana As Boolean, spaces As Boolean) As String
    
    Dim result As String
    result = inputText
    
    If Len(result) = 0 Then
        ConvertText = result
        Exit Function
    End If
    
    If direction = 1 Then
        ' 全角→半角変換
        If alphaNumeric Then
            result = ConvertAlphaNumericToHankaku(result)
        End If
        
        If symbols Then
            result = ConvertSymbolsToHankaku(result)
        End If
        
        If katakana Then
            result = StrConv(result, vbNarrow + vbKatakana, &H411)
        End If
        
        If spaces Then
            result = Replace(result, "　", " ") ' 全角スペース→半角スペース
        End If
        
    Else
        ' 半角→全角変換
        If alphaNumeric Then
            result = ConvertAlphaNumericToZenkaku(result)
        End If
        
        If symbols Then
            result = ConvertSymbolsToZenkaku(result)
        End If
        
        If katakana Then
            result = StrConv(result, vbWide + vbKatakana, &H411)
        End If
        
        If spaces Then
            result = Replace(result, " ", "　") ' 半角スペース→全角スペース
        End If
    End If
    
    ConvertText = result
End Function

Private Function ConvertAlphaNumericToHankaku(inputText As String) As String
    Dim result As String
    Dim i As Integer
    Dim char As String
    Dim convertedChar As String
    
    result = ""
    For i = 1 To Len(inputText)
        char = Mid(inputText, i, 1)
        
        ' 全角英数字の範囲をチェックして変換
        Select Case AscW(char)
            Case &HFF10 To &HFF19 ' 全角数字 ０-９
                convertedChar = Chr(AscW(char) - &HFF10 + Asc("0"))
            Case &HFF21 To &HFF3A ' 全角英字 Ａ-Ｚ、記号
                convertedChar = Chr(AscW(char) - &HFF00 + &H20)
            Case &HFF41 To &HFF5A ' 全角英字 ａ-ｚ、記号
                convertedChar = Chr(AscW(char) - &HFF00 + &H20)
            Case Else
                convertedChar = char
        End Select
        
        result = result & convertedChar
    Next i
    
    ConvertAlphaNumericToHankaku = result
End Function

Private Function ConvertAlphaNumericToZenkaku(inputText As String) As String
    Dim result As String
    Dim i As Integer
    Dim char As String
    Dim convertedChar As String
    
    result = ""
    For i = 1 To Len(inputText)
        char = Mid(inputText, i, 1)
        
        ' 半角英数字の範囲をチェックして変換
        Select Case Asc(char)
            Case 48 To 57 ' 半角数字 0-9
                convertedChar = ChrW(AscW(char) + &HFF00 - &H20)
            Case 65 To 90 ' 半角英字 A-Z
                convertedChar = ChrW(AscW(char) + &HFF00 - &H20)
            Case 97 To 122 ' 半角英字 a-z
                convertedChar = ChrW(AscW(char) + &HFF00 - &H20)
            Case Else
                convertedChar = char
        End Select
        
        result = result & convertedChar
    Next i
    
    ConvertAlphaNumericToZenkaku = result
End Function

Private Function ConvertSymbolsToHankaku(inputText As String) As String
    Dim result As String
    result = inputText
    
    ' よく使用される記号の変換マップ
    result = Replace(result, "！", "!")
    result = Replace(result, "？", "?")
    result = Replace(result, "．", ".")
    result = Replace(result, "，", ",")
    result = Replace(result, "：", ":")
    result = Replace(result, "；", ";")
    result = Replace(result, "（", "(")
    result = Replace(result, "）", ")")
    result = Replace(result, "［", "[")
    result = Replace(result, "］", "]")
    result = Replace(result, "｛", "{")
    result = Replace(result, "｝", "}")
    result = Replace(result, "「", """")
    result = Replace(result, "」", """")
    result = Replace(result, "＋", "+")
    result = Replace(result, "－", "-")
    result = Replace(result, "＝", "=")
    result = Replace(result, "＜", "<")
    result = Replace(result, "＞", ">")
    result = Replace(result, "％", "%")
    result = Replace(result, "＆", "&")
    result = Replace(result, "＃", "#")
    result = Replace(result, "＄", "$")
    result = Replace(result, "＠", "@")
    result = Replace(result, "＊", "*")
    result = Replace(result, "／", "/")
    result = Replace(result, "￥", "\")
    
    ConvertSymbolsToHankaku = result
End Function

Private Function ConvertSymbolsToZenkaku(inputText As String) As String
    Dim result As String
    result = inputText
    
    ' よく使用される記号の変換マップ（半角→全角）
    result = Replace(result, "!", "！")
    result = Replace(result, "?", "？")
    result = Replace(result, ".", "．")
    result = Replace(result, ",", "，")
    result = Replace(result, ":", "：")
    result = Replace(result, ";", "；")
    result = Replace(result, "(", "（")
    result = Replace(result, ")", "）")
    result = Replace(result, "[", "［")
    result = Replace(result, "]", "］")
    result = Replace(result, "{", "｛")
    result = Replace(result, "}", "｝")
    result = Replace(result, """", "「")
    result = Replace(result, "+", "＋")
    result = Replace(result, "-", "－")
    result = Replace(result, "=", "＝")
    result = Replace(result, "<", "＜")
    result = Replace(result, ">", "＞")
    result = Replace(result, "%", "％")
    result = Replace(result, "&", "＆")
    result = Replace(result, "#", "＃")
    result = Replace(result, "$", "＄")
    result = Replace(result, "@", "＠")
    result = Replace(result, "*", "＊")
    result = Replace(result, "/", "／")
    result = Replace(result, "\", "￥")
    
    ConvertSymbolsToZenkaku = result
End Function

Private Function ShouldProcessCell(cell As Range, IncludeFormulas As Boolean) As Boolean
    ' 空セルやエラーセルはスキップ
    If IsEmpty(cell) Or IsError(cell) Then
        ShouldProcessCell = False
        Exit Function
    End If
    
    ' 数式セルの処理判定
    If cell.HasFormula And Not IncludeFormulas Then
        ShouldProcessCell = False
        Exit Function
    End If
    
    ShouldProcessCell = True
End Function

Sub QuickZenkakuToHankaku()
    ' 全角→半角（英数字＋記号）
    Call QuickConvert(1, True, True, False, False)
End Sub

Sub QuickHankakuToZenkaku()
    ' 半角→全角（英数字＋記号）
    Call QuickConvert(2, True, True, False, False)
End Sub

Sub QuickZenkakuToHankakuAll()
    ' 全角→半角（全て）
    Call QuickConvert(1, True, True, True, True)
End Sub

Sub QuickHankakuToZenkakuAll()
    ' 半角→全角（全て）
    Call QuickConvert(2, True, True, True, True)
End Sub

Private Sub QuickConvert(direction As Integer, alphaNum As Boolean, _
                       symbols As Boolean, katakana As Boolean, spaces As Boolean)
    On Error GoTo QuickErrorHandler
    
    Dim cell As Range
    Dim changedCount As Long
    Dim rng As Range
    
    Set rng = selection
    If rng Is Nothing Or rng.count = 0 Then
        MsgBox "セルを選択してください。", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    For Each cell In rng
        If ShouldProcessCell(cell, False) Then
            Dim originalValue As Variant
            Dim convertedValue As String
            
            originalValue = cell.value
            
            ' 数値の場合は文字列として処理
            If IsNumeric(originalValue) And Not IsEmpty(originalValue) Then
                cell.NumberFormat = "@"
                cell.value = CStr(originalValue)
                originalValue = CStr(originalValue)
            End If
            
            convertedValue = ConvertText(CStr(originalValue), direction, alphaNum, symbols, katakana, spaces)
            
            If CStr(originalValue) <> convertedValue Then
                cell.NumberFormat = "@"
                cell.value = convertedValue
                
                ' 全角→半角で数値のみの場合は数値形式に戻す
                If direction = 1 And IsNumericOnly(convertedValue) Then
                    On Error Resume Next
                    Dim numValue As Double
                    numValue = CDbl(convertedValue)
                    If Err.Number = 0 Then
                        cell.NumberFormat = "General"
                        cell.value = numValue
                    End If
                    On Error GoTo QuickErrorHandler
                End If
                
                changedCount = changedCount + 1
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Dim directionText As String
    directionText = IIf(direction = 1, "全角→半角", "半角→全角")
    MsgBox changedCount & " セルを" & directionText & "変換しました。", vbInformation
    
    Exit Sub
    
QuickErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "エラーが発生しました: " & Err.description, vbCritical
End Sub

