VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmZenkakuHankaku 
   Caption         =   "�S�p���p�ϊ�"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4350
   OleObjectBlob   =   "frmZenkakuHankaku.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmZenkakuHankaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' �S�p���p�ϊ��}�N���i�G���[�C���Łj
' �쐬���F2025/08/19
' �@�\�F�I��͈͂̑S�p�E���p�����𓝈�ϊ�
' �C���F�t�H�[���I�����G���[�Ή��A���l�`�������@�\�ǉ�
'==============================================================================

'==============================================================================
' ���[�U�[�t�H�[��: frmZenkakuHankaku
' �S�p���p�ϊ��̐ݒ���s�����߂̃t�H�[��
'==============================================================================

' �t�H�[�����x���ϐ��i�t�H�[�����W���[���ɋL�q�j
Private m_ProcessExecuted As Boolean
Private m_ConversionDirection As Integer    ' 1: �S�p�����p, 2: ���p���S�p
Private m_ConvertAlphaNumeric As Boolean
Private m_ConvertSymbols As Boolean
Private m_ConvertKatakana As Boolean
Private m_ConvertSpaces As Boolean
Private m_IncludeFormulas As Boolean

' �v���p�e�B�i���C���}�N������Q�Ɨp�j
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

' �t�H�[���������i�t�H�[�����W���[���ɋL�q�j
Private Sub UserForm_Initialize()
    On Error Resume Next
    
    ' �f�t�H���g�l�ݒ�
    m_ProcessExecuted = False
    m_ConversionDirection = 1 ' �S�p�����p
    m_ConvertAlphaNumeric = True
    m_ConvertSymbols = True
    m_ConvertKatakana = False
    m_ConvertSpaces = False
    m_IncludeFormulas = False
    
    ' �R���g���[���̃f�t�H���g�l�ݒ�
    optZenToHan.value = True
    chkAlphaNumeric.value = True
    chkSymbols.value = False
    chkKatakana.value = False
    chkSpaces.value = False
    chkFormulas.value = False
    
    On Error GoTo 0
End Sub

' ���s�{�^���N���b�N�i�t�H�[�����W���[���ɋL�q�j
Private Sub btnOK_Click()
    On Error Resume Next
    
    ' �ݒ��ϐ��ɕۑ�
    If optZenToHan.value Then
        m_ConversionDirection = 1 ' �S�p�����p
    Else
        m_ConversionDirection = 2 ' ���p���S�p
    End If
    
    ' �ϊ��Ώۂ̎擾
    m_ConvertAlphaNumeric = chkAlphaNumeric.value
    m_ConvertSymbols = chkSymbols.value
    m_ConvertKatakana = chkKatakana.value
    m_ConvertSpaces = chkSpaces.value
    m_IncludeFormulas = chkFormulas.value
    
    ' �Œ�1�̕ϊ��Ώۂ��I������Ă��邩�`�F�b�N
    If Not (m_ConvertAlphaNumeric Or m_ConvertSymbols Or m_ConvertKatakana Or m_ConvertSpaces) Then
        MsgBox "�ϊ��Ώۂ�1�ȏ�I�����Ă��������B", vbExclamation, "�ݒ�G���["
        On Error GoTo 0
        Exit Sub
    End If
    
    ' �ŏI�m�F
    Dim confirmMsg As String
    Dim directionText As String
    directionText = IIf(m_ConversionDirection = 1, "�S�p�����p", "���p���S�p")
    
    Dim targetText As String
    If m_ConvertAlphaNumeric Then targetText = targetText & "�p���� "
    If m_ConvertSymbols Then targetText = targetText & "�L�� "
    If m_ConvertKatakana Then targetText = targetText & "�J�^�J�i "
    If m_ConvertSpaces Then targetText = targetText & "�X�y�[�X "
    
    confirmMsg = "�ȉ��̐ݒ�ŕϊ������s���܂��F" & vbCrLf & vbCrLf & _
                "�ϊ�����: " & directionText & vbCrLf & _
                "�ϊ��Ώ�: " & Trim(targetText) & vbCrLf & _
                "�����Z��: " & IIf(m_IncludeFormulas, "�܂�", "���O") & vbCrLf & vbCrLf & _
                "���s���܂����H"
    
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "�ϊ����s�m�F") = vbYes Then
        m_ProcessExecuted = True
        Me.Hide
    End If
    
    On Error GoTo 0
End Sub

' �L�����Z���{�^���N���b�N�i�t�H�[�����W���[���ɋL�q�j
Private Sub btnCancel_Click()
    On Error Resume Next
    m_ProcessExecuted = False
    Me.Hide
    On Error GoTo 0
End Sub

' �~�{�^���i����{�^���j�Ή�
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    ' �E��́~�{�^���������ꂽ�ꍇ
    If CloseMode = 0 Then  ' vbFormControlMenu (�~�{�^��)
        Cancel = True  ' �ʏ�̕��鏈�����L�����Z��
        m_SelectedOption = 0  ' �L�����Z��������
        Me.Hide  ' �t�H�[�����\���ɂ���
    End If
    ' �G���[�𖳎����ăt�H�[�����m���ɕ���
    On Error GoTo 0
End Sub
