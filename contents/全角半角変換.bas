Attribute VB_Name = "�S�p���p�ϊ�"
Sub ZenkakuHankakuConverter()

    Dim ws As Worksheet
    Dim rng As Range
    Dim frm As Object
    
    ' �G���[�n���h�����O
    On Error GoTo ErrorHandler
    
    ' �A�N�e�B�u�ȃ��[�N�V�[�g�擾
    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "�A�N�e�B�u�ȃ��[�N�V�[�g������܂���B", vbExclamation
        Exit Sub
    End If
    
    ' �V�[�g���ی삳��Ă��邩�`�F�b�N
    If ws.ProtectContents Then
        MsgBox "�V�[�g�u" & ws.Name & "�v���ی삳��Ă��܂��B�ی���������Ă�����s���Ă��������B", vbExclamation
        Exit Sub
    End If
    
    ' �I��͈͂̊m�F�Ɛݒ�
    Set rng = selection
    If rng Is Nothing Then
        MsgBox "�͈͂��I������Ă��܂���B", vbExclamation
        Exit Sub
    End If
    
    ' �P��̋�Z�����I������Ă���ꍇ
    If rng.count = 1 And IsEmpty(rng) Then
        If MsgBox("�I������Ă���Z������ł��B" & vbCrLf & _
                 "�g�p����Ă���S�͈͂�Ώۂɂ��܂����H", _
                 vbYesNo + vbQuestion) = vbYes Then
            Set rng = ws.UsedRange
        Else
            Exit Sub
        End If
    End If
    
    ' �t�H�[����\�����Đݒ�擾
    Dim useForm As Boolean
    useForm = True
    
    ' �t�H�[���̍쐬�����s
    On Error Resume Next
    Set frm = VBA.UserForms.Add("frmZenkakuHankaku")
    If Err.Number <> 0 Or frm Is Nothing Then
        useForm = False
    End If
    On Error GoTo 0
    
    If useForm Then
        ' �t�H�[�������p�\�ȏꍇ
        On Error Resume Next
        frm.Show
        If Err.Number <> 0 Then
            ' �t�H�[���\���ŃG���[�����������ꍇ
            useForm = False
            Unload frm
            Set frm = Nothing
        End If
        On Error GoTo 0
        
        If useForm Then
            If Not frm.ProcessExecuted Then
                ' �L�����Z���܂��́~�ŕ���ꂽ�ꍇ�͏I��
                Unload frm
                Set frm = Nothing
                Exit Sub
            End If
            
            ' �ݒ���擾
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
            
            ' �t�H�[�������
            Unload frm
            Set frm = Nothing
            
            ' �ϊ����s
            Call ExecuteConversion(rng, convDirection, convAlphaNumeric, convSymbols, convKatakana, convSpaces, IncludeFormulas)
        End If
    End If
    
    If Not useForm Then
        ' �t�H�[�������p�ł��Ȃ��ꍇ�͊ȈՔł��g�p
        If Not GetSettingsSimple(rng) Then
            Exit Sub
        End If
    End If
    
ErrorHandler:
    ' �G���[�ԍ�0�̏ꍇ�͉������Ȃ�
    If Err.Number = 0 Then
        Exit Sub
    End If
    
    ' �t�H�[���̃N���[���A�b�v
    On Error Resume Next
    If Not frm Is Nothing Then
        Unload frm
        Set frm = Nothing
    End If
    On Error GoTo 0
    
    ' �G���[�ԍ��Ɠ��e���L���ȏꍇ�̂݃��b�Z�[�W��\��
    If Err.Number <> 0 Then
        MsgBox "�G���[���������܂����B" & vbCrLf & _
               "�G���[�ԍ�: " & Err.Number & vbCrLf & _
               "�G���[���e: " & Err.description, vbCritical
    End If
End Sub

Private Function GetSettingsSimple(rng As Range) As Boolean
    Dim direction As Integer
    Dim alphaNumeric As Boolean, symbols As Boolean, katakana As Boolean, spaces As Boolean
    Dim IncludeFormulas As Boolean
    
    On Error Resume Next
    
    ' �ϊ������I��
    If MsgBox("�ϊ�������I�����Ă�������" & vbCrLf & vbCrLf & _
             "�u�͂��v: �S�p �� ���p" & vbCrLf & _
             "�u�������v: ���p �� �S�p", _
             vbYesNo + vbQuestion, "�ϊ������I��") = vbYes Then
        direction = 1
    Else
        direction = 2
    End If
    
    ' �ϊ��ΏۑI��
    alphaNumeric = (MsgBox("�p������ϊ����܂����H", vbYesNo + vbQuestion) = vbYes)
    symbols = (MsgBox("�L����ϊ����܂����H", vbYesNo + vbQuestion) = vbYes)
    katakana = (MsgBox("�J�^�J�i��ϊ����܂����H", vbYesNo + vbQuestion) = vbYes)
    spaces = (MsgBox("�X�y�[�X��ϊ����܂����H", vbYesNo + vbQuestion) = vbYes)
    IncludeFormulas = (MsgBox("�����Z�����������܂����H", vbYesNo + vbQuestion) = vbYes)
    
    ' �Œ�1�̕ϊ��Ώۂ��I������Ă��邩�`�F�b�N
    If Not (alphaNumeric Or symbols Or katakana Or spaces) Then
        MsgBox "�ϊ��Ώۂ��I������Ă��܂���B", vbExclamation
        GetSettingsSimple = False
        On Error GoTo 0
        Exit Function
    End If
    
    On Error GoTo 0
    
    ' �ϊ����s
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
    
    ' ��ʍX�V���~���ď������x����
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' ���C������
    For Each cell In rng
        If ShouldProcessCell(cell, IncludeFormulas) Then
            originalValue = cell.value
            Dim originalFormat As String
            originalFormat = cell.NumberFormat
            
            ' ���l�̏ꍇ�͕�����Ƃ��ď���
            If IsNumeric(originalValue) And Not IsEmpty(originalValue) Then
                ' �Z���̕\���`���𕶎���ɕύX
                cell.NumberFormat = "@"
                ' �l�𕶎���Ƃ��čĐݒ�
                cell.value = CStr(originalValue)
                originalValue = CStr(originalValue)
            End If
            
            convertedValue = ConvertText(CStr(originalValue), direction, alphaNumeric, symbols, katakana, spaces)
            
            If CStr(originalValue) <> convertedValue Then
                ' �ϊ���̒l��ݒ�
                cell.NumberFormat = "@"
                cell.value = convertedValue
                
                ' �S�p�����p�Ő��l�݂̂̏ꍇ�͐��l�`���ɖ߂�
                If direction = 1 And IsNumericOnly(convertedValue) Then
                    ' ���l�Ƃ��ĔF���ł���ꍇ�͐��l�`���ɖ߂�
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
    
    ' �ݒ�����ɖ߂�
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' �������ʂ��
    Dim resultMsg As String
    Dim processingTime As Double
    processingTime = Timer - startTime
    
    Dim directionText As String
    directionText = IIf(direction = 1, "�S�p�����p", "���p���S�p")
    
    resultMsg = "�S�p���p���ꏈ�����������܂����B" & vbCrLf & vbCrLf & _
               "��������:" & vbCrLf & _
               "? �ϊ�����: " & directionText & vbCrLf & _
               "? �����Z����: " & Format(processedCount, "#,##0") & vbCrLf & _
               "? �ύX�Z����: " & Format(changedCount, "#,##0") & vbCrLf & _
               "? ��������: " & Format(processingTime, "0.00") & "�b"
    
    MsgBox resultMsg, vbInformation, "��������"
    
    Exit Sub
    
ConversionError:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "�ϊ��������ɃG���[���������܂����B" & vbCrLf & _
           "�G���[���e: " & Err.description, vbCritical
End Sub

Private Function IsNumericOnly(value As String) As Boolean
    On Error Resume Next
    
    ' �󕶎���͐��l�ł͂Ȃ�
    If Len(Trim(value)) = 0 Then
        IsNumericOnly = False
        Exit Function
    End If
    
    ' ���l�Ƃ��ĕϊ��ł��邩�`�F�b�N
    Dim testValue As Double
    testValue = CDbl(value)
    
    If Err.Number = 0 Then
        ' ����ɕ����񂪐����A�����_�A�����݂̂ō\������Ă��邩�`�F�b�N
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
        ' �S�p�����p�ϊ�
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
            result = Replace(result, "�@", " ") ' �S�p�X�y�[�X�����p�X�y�[�X
        End If
        
    Else
        ' ���p���S�p�ϊ�
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
            result = Replace(result, " ", "�@") ' ���p�X�y�[�X���S�p�X�y�[�X
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
        
        ' �S�p�p�����͈̔͂��`�F�b�N���ĕϊ�
        Select Case AscW(char)
            Case &HFF10 To &HFF19 ' �S�p���� �O-�X
                convertedChar = Chr(AscW(char) - &HFF10 + Asc("0"))
            Case &HFF21 To &HFF3A ' �S�p�p�� �`-�y�A�L��
                convertedChar = Chr(AscW(char) - &HFF00 + &H20)
            Case &HFF41 To &HFF5A ' �S�p�p�� ��-���A�L��
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
        
        ' ���p�p�����͈̔͂��`�F�b�N���ĕϊ�
        Select Case Asc(char)
            Case 48 To 57 ' ���p���� 0-9
                convertedChar = ChrW(AscW(char) + &HFF00 - &H20)
            Case 65 To 90 ' ���p�p�� A-Z
                convertedChar = ChrW(AscW(char) + &HFF00 - &H20)
            Case 97 To 122 ' ���p�p�� a-z
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
    
    ' �悭�g�p�����L���̕ϊ��}�b�v
    result = Replace(result, "�I", "!")
    result = Replace(result, "�H", "?")
    result = Replace(result, "�D", ".")
    result = Replace(result, "�C", ",")
    result = Replace(result, "�F", ":")
    result = Replace(result, "�G", ";")
    result = Replace(result, "�i", "(")
    result = Replace(result, "�j", ")")
    result = Replace(result, "�m", "[")
    result = Replace(result, "�n", "]")
    result = Replace(result, "�o", "{")
    result = Replace(result, "�p", "}")
    result = Replace(result, "�u", """")
    result = Replace(result, "�v", """")
    result = Replace(result, "�{", "+")
    result = Replace(result, "�|", "-")
    result = Replace(result, "��", "=")
    result = Replace(result, "��", "<")
    result = Replace(result, "��", ">")
    result = Replace(result, "��", "%")
    result = Replace(result, "��", "&")
    result = Replace(result, "��", "#")
    result = Replace(result, "��", "$")
    result = Replace(result, "��", "@")
    result = Replace(result, "��", "*")
    result = Replace(result, "�^", "/")
    result = Replace(result, "��", "\")
    
    ConvertSymbolsToHankaku = result
End Function

Private Function ConvertSymbolsToZenkaku(inputText As String) As String
    Dim result As String
    result = inputText
    
    ' �悭�g�p�����L���̕ϊ��}�b�v�i���p���S�p�j
    result = Replace(result, "!", "�I")
    result = Replace(result, "?", "�H")
    result = Replace(result, ".", "�D")
    result = Replace(result, ",", "�C")
    result = Replace(result, ":", "�F")
    result = Replace(result, ";", "�G")
    result = Replace(result, "(", "�i")
    result = Replace(result, ")", "�j")
    result = Replace(result, "[", "�m")
    result = Replace(result, "]", "�n")
    result = Replace(result, "{", "�o")
    result = Replace(result, "}", "�p")
    result = Replace(result, """", "�u")
    result = Replace(result, "+", "�{")
    result = Replace(result, "-", "�|")
    result = Replace(result, "=", "��")
    result = Replace(result, "<", "��")
    result = Replace(result, ">", "��")
    result = Replace(result, "%", "��")
    result = Replace(result, "&", "��")
    result = Replace(result, "#", "��")
    result = Replace(result, "$", "��")
    result = Replace(result, "@", "��")
    result = Replace(result, "*", "��")
    result = Replace(result, "/", "�^")
    result = Replace(result, "\", "��")
    
    ConvertSymbolsToZenkaku = result
End Function

Private Function ShouldProcessCell(cell As Range, IncludeFormulas As Boolean) As Boolean
    ' ��Z����G���[�Z���̓X�L�b�v
    If IsEmpty(cell) Or IsError(cell) Then
        ShouldProcessCell = False
        Exit Function
    End If
    
    ' �����Z���̏�������
    If cell.HasFormula And Not IncludeFormulas Then
        ShouldProcessCell = False
        Exit Function
    End If
    
    ShouldProcessCell = True
End Function

Sub QuickZenkakuToHankaku()
    ' �S�p�����p�i�p�����{�L���j
    Call QuickConvert(1, True, True, False, False)
End Sub

Sub QuickHankakuToZenkaku()
    ' ���p���S�p�i�p�����{�L���j
    Call QuickConvert(2, True, True, False, False)
End Sub

Sub QuickZenkakuToHankakuAll()
    ' �S�p�����p�i�S�āj
    Call QuickConvert(1, True, True, True, True)
End Sub

Sub QuickHankakuToZenkakuAll()
    ' ���p���S�p�i�S�āj
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
        MsgBox "�Z����I�����Ă��������B", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    For Each cell In rng
        If ShouldProcessCell(cell, False) Then
            Dim originalValue As Variant
            Dim convertedValue As String
            
            originalValue = cell.value
            
            ' ���l�̏ꍇ�͕�����Ƃ��ď���
            If IsNumeric(originalValue) And Not IsEmpty(originalValue) Then
                cell.NumberFormat = "@"
                cell.value = CStr(originalValue)
                originalValue = CStr(originalValue)
            End If
            
            convertedValue = ConvertText(CStr(originalValue), direction, alphaNum, symbols, katakana, spaces)
            
            If CStr(originalValue) <> convertedValue Then
                cell.NumberFormat = "@"
                cell.value = convertedValue
                
                ' �S�p�����p�Ő��l�݂̂̏ꍇ�͐��l�`���ɖ߂�
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
    directionText = IIf(direction = 1, "�S�p�����p", "���p���S�p")
    MsgBox changedCount & " �Z����" & directionText & "�ϊ����܂����B", vbInformation
    
    Exit Sub
    
QuickErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "�G���[���������܂���: " & Err.description, vbCritical
End Sub

