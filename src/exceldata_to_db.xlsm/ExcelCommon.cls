VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//----------------------------------------------------------------------------
'// ExcelCommon�N���X
'//----------------------------------------------------------------------------


'// ���֐����R�����g�p�̃t�H�[�}�b�g
'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F
'// ����    �F
'// �߂�l  �F
'// ���l    �F
'//----------------------------------------------------------------------------

Option Explicit

Private book As Workbook
Private opened_book As Workbook

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�R���X�g���N�^
'// ����    �F
'// �߂�l  �F
'// ���l    �F
'//----------------------------------------------------------------------------
Private Sub Class_Initialize()

End Sub

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F���[�N�u�b�N�̃Z�b�g�B�C���X�^���X����Ɏg�p����B
'// ����    �F
'// �߂�l  �F
'// ���l    �F
'//----------------------------------------------------------------------------
Public Sub set_workbook(ByVal workbook_ As Workbook)

    Set book = workbook_

End Sub

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�w�肵���V�[�g�ƃZ���͈͂̃f�[�^���N���A
'// ����    �F
'// �߂�l  �F
'// ���l    �F
'//----------------------------------------------------------------------------
Public Sub data_clear(ByVal sheet As String, start_row As Long, start_col As Long, end_row As Long, end_col As Long)
    
    book.Worksheets(sheet).Activate
    book.Worksheets(sheet).Range(Cells(start_row, start_col), Cells(end_row, end_col)).Clear

End Sub

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�w�肵���V�[�g�ƃZ���͈͂̃f�[�^���N���A
'// ����    �F
'// �߂�l  �F
'// ���l    �F�w�肵���s�񂩂�E���A�����ɑ����ő�͈̔͂��v�Z���ăN���A����B
'//----------------------------------------------------------------------------
Public Sub data_clear2(ByVal sheet As String, start_row As Long, start_col As Long)
    
    Dim end_row As Long
    Dim end_col As Long
    
    end_row = Me.get_max_row_below(sheet, start_row, start_col)
    end_col = Me.get_max_col_right(sheet, start_row, start_col)
    
    book.Worksheets(sheet).Activate
    book.Worksheets(sheet).Range(Cells(start_row, start_col), Cells(end_row, end_col)).Clear

End Sub


'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�w�肵���Z�����牺���ɑ����ő�̃Z���s���擾
'// ����    �F
'// �߂�l  �F
'// ���l    �F
'//----------------------------------------------------------------------------
Public Function get_max_row_below(ByVal sheet As String, row As Long, col As Long) As Long

    If book Is Nothing Then
        MsgBox "Class ExcelCommon �Ƀ��[�N�u�b�N���Z�b�g����Ă��܂���B"
        Exit Function
    End If

    Dim i As Long
    i = 1
    Do While book.Worksheets(sheet).Cells(row + i, col) <> ""
        i = i + 1
    Loop
    
    get_max_row_below = row + i - 1

End Function


'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�w�肵���Z������E���ɑ����ő�̃Z������擾
'// ����    �F
'// �߂�l  �F
'// ���l    �F
'//----------------------------------------------------------------------------
Public Function get_max_col_right(ByVal sheet As String, row As Long, col As Long) As Long

    If book Is Nothing Then
        MsgBox "Class ExcelCommon �Ƀ��[�N�u�b�N���Z�b�g����Ă��܂���B"
        Exit Function
    End If

    Dim i As Long
    i = 1
    Do While book.Worksheets(sheet).Cells(row, col + i) <> ""
        i = i + 1
    Loop
    
    get_max_col_right = col + i - 1

End Function

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�w�肵���t�@�C�����J��
'// ����    �F
'// �߂�l  �F
'// ���l    �F�t�@�C�����̓t�H���_���܂߂ēn��
'//----------------------------------------------------------------------------
Public Sub open_file(ByVal file_name As String)

    Set opened_book = Workbooks.Open(file_name)
    Debug.Print file_name & "���J���܂����B"

End Sub

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�J���Ă����t�@�C�������
'// ����    �F
'// �߂�l  �F
'// ���l    �F
'//----------------------------------------------------------------------------
Public Sub close_opened_file(ByVal display_alerts As Boolean)

    Application.DisplayAlerts = display_alerts
    opened_book.Close
    Debug.Print "�J���Ă����t�@�C������܂����B"
    
    Application.DisplayAlerts = True

End Sub