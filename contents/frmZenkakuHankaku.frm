VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmZenkakuHankaku 
   Caption         =   "全角半角変換"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4350
   OleObjectBlob   =   "frmZenkakuHankaku.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmZenkakuHankaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' 全角半角変換マクロ（エラー修正版）
' 作成日：2025/08/19
' 機能：選択範囲の全角・半角文字を統一変換
' 修正：フォーム終了時エラー対応、数値形式復元機能追加
'==============================================================================

'==============================================================================
' ユーザーフォーム: frmZenkakuHankaku
' 全角半角変換の設定を行うためのフォーム
'==============================================================================

' フォームレベル変数（フォームモジュールに記述）
Private m_ProcessExecuted As Boolean
Private m_ConversionDirection As Integer    ' 1: 全角→半角, 2: 半角→全角
Private m_ConvertAlphaNumeric As Boolean
Private m_ConvertSymbols As Boolean
Private m_ConvertKatakana As Boolean
Private m_ConvertSpaces As Boolean
Private m_IncludeFormulas As Boolean

' プロパティ（メインマクロから参照用）
Public Property Get ProcessExecuted() As Boolean
    ProcessExecuted = m_ProcessExecuted
End Property

Public Property Get ConversionDirection() As Integer
    ConversionDirection = m_ConversionDirection
End Property

Public Property Get ConvertAlphaNumeric() As Boolean
    ConvertAlphaNumeric = m_ConvertAlphaNumeric
End Property

Public Property Get ConvertSymbols() As Boolean
    ConvertSymbols = m_ConvertSymbols
End Property

Public Property Get ConvertKatakana() As Boolean
    ConvertKatakana = m_ConvertKatakana
End Property

Public Property Get ConvertSpaces() As Boolean
    ConvertSpaces = m_ConvertSpaces
End Property

Public Property Get IncludeFormulas() As Boolean
    IncludeFormulas = m_IncludeFormulas
End Property

' フォーム初期化（フォームモジュールに記述）
Private Sub UserForm_Initialize()
    On Error Resume Next
    
    ' デフォルト値設定
    m_ProcessExecuted = False
    m_ConversionDirection = 1 ' 全角→半角
    m_ConvertAlphaNumeric = True
    m_ConvertSymbols = True
    m_ConvertKatakana = False
    m_ConvertSpaces = False
    m_IncludeFormulas = False
    
    ' コントロールのデフォルト値設定
    optZenToHan.value = True
    chkAlphaNumeric.value = True
    chkSymbols.value = False
    chkKatakana.value = False
    chkSpaces.value = False
    chkFormulas.value = False
    
    On Error GoTo 0
End Sub

' 実行ボタンクリック（フォームモジュールに記述）
Private Sub btnOK_Click()
    On Error Resume Next
    
    ' 設定を変数に保存
    If optZenToHan.value Then
        m_ConversionDirection = 1 ' 全角→半角
    Else
        m_ConversionDirection = 2 ' 半角→全角
    End If
    
    ' 変換対象の取得
    m_ConvertAlphaNumeric = chkAlphaNumeric.value
    m_ConvertSymbols = chkSymbols.value
    m_ConvertKatakana = chkKatakana.value
    m_ConvertSpaces = chkSpaces.value
    m_IncludeFormulas = chkFormulas.value
    
    ' 最低1つの変換対象が選択されているかチェック
    If Not (m_ConvertAlphaNumeric Or m_ConvertSymbols Or m_ConvertKatakana Or m_ConvertSpaces) Then
        MsgBox "変換対象を1つ以上選択してください。", vbExclamation, "設定エラー"
        On Error GoTo 0
        Exit Sub
    End If
    
    ' 最終確認
    Dim confirmMsg As String
    Dim directionText As String
    directionText = IIf(m_ConversionDirection = 1, "全角→半角", "半角→全角")
    
    Dim targetText As String
    If m_ConvertAlphaNumeric Then targetText = targetText & "英数字 "
    If m_ConvertSymbols Then targetText = targetText & "記号 "
    If m_ConvertKatakana Then targetText = targetText & "カタカナ "
    If m_ConvertSpaces Then targetText = targetText & "スペース "
    
    confirmMsg = "以下の設定で変換を実行します：" & vbCrLf & vbCrLf & _
                "変換方向: " & directionText & vbCrLf & _
                "変換対象: " & Trim(targetText) & vbCrLf & _
                "数式セル: " & IIf(m_IncludeFormulas, "含む", "除外") & vbCrLf & vbCrLf & _
                "実行しますか？"
    
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "変換実行確認") = vbYes Then
        m_ProcessExecuted = True
        Me.Hide
    End If
    
    On Error GoTo 0
End Sub

' キャンセルボタンクリック（フォームモジュールに記述）
Private Sub btnCancel_Click()
    On Error Resume Next
    m_ProcessExecuted = False
    Me.Hide
    On Error GoTo 0
End Sub

' ×ボタン（閉じるボタン）対応
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    ' 右上の×ボタンが押された場合
    If CloseMode = 0 Then  ' vbFormControlMenu (×ボタン)
        Cancel = True  ' 通常の閉じる処理をキャンセル
        m_SelectedOption = 0  ' キャンセルを示す
        Me.Hide  ' フォームを非表示にする
    End If
    ' エラーを無視してフォームを確実に閉じる
    On Error GoTo 0
End Sub
