VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelDataGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'//----------------------------------------------------------------------------
'// ExcelDataGetter�N���X
'//----------------------------------------------------------------------------

Option Explicit

'//----------------------------------------------------------------------------
'// �y�g�p��z
'//  �@�C���X�^���X����
'//�@�Aset_info�F�f�[�^��ǂݍ���Excel�����C���v�b�g
'//  [�f�[�^�擾�Ɏ��Ԃ��|����Ȃ�΁A���̊ԂɓǍ��t�@�C�����J������������i�ʓr�j]
'//�@�Bdefine_data_rows�F�f�[�^��ǂݍ��ލs�̊J�n����I���s������
'//�@�Cset_data_array�F����̗�ɑ΂���f�[�^���ꎞ�I�Ɋi�[
'//�@�Dpaste_data_to_excel�F�w�肵��Excel�t�@�C���̃Z���Ɋi�[�f�[�^��\��t��
'//----------------------------------------------------------------------------


Private folder_name As String
Private file_name As String
Private sheet_name As String
Private target As String            '// �t�H���_�� + "[" + �t�@�C���� + "]" + �V�[�g�� + "!"

Private data_first_row As Long      '// ��Excel�̃f�[�^���i�[����Ă���ŏ��̍s
Private data_last_row As Long       '// ��Excel�̃f�[�^���i�[����Ă���Ō�̍s
Private data() As Variant           '// ��Excel�̃f�[�^���i�[����z��

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�R���X�g���N�^
'// ����    �F
'// �߂�l  �F
'// ���l    �F
'//----------------------------------------------------------------------------
Private Sub Class_Initialize()

    folder_name = ""
    file_name = ""
    sheet_name = ""
    target = ""

    data_first_row = -1
    data_last_row = -1
    
    ReDim data(0)

End Sub

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�f�[�^��ǂݍ��ޑΏ�Excel�̃t�H���_�A�t�@�C�����A�V�[�g����ݒ�
'// ����    �F
'// �߂�l  �F
'// ���l    �F
'//----------------------------------------------------------------------------
Public Sub set_info(ByVal folder_name_ As String, file_name_ As String, sheet_name_ As String)

    Dim buf As String
    Dim open_file_name As String

    folder_name = folder_name_
    file_name = file_name_
    sheet_name = sheet_name_

    ''�Ώۃu�b�N��I�����܂�
    '// open_file_name = Application.GetOpenFilename("Microsoft Excel �u�b�N,*.xls")
    open_file_name = folder_name_ & "\" & file_name_
    
    '// �t�@�C�������݂��Ȃ���ΏI��
    If Dir(open_file_name) = "" Then
        MsgBox (open_file_name & " �����݂��܂���B" & Chr(13) _
            & "�t�H���_���ƃt�@�C�������m�F���ĉ������B")
        Exit Sub
    End If
    
    ''�t�@�C������[]��t����
    open_file_name = Replace(open_file_name, Dir(open_file_name), "[" & Dir(open_file_name) & "]")
    'Debug.Print (open_file_name)
    
    ''�Ώۃ��[�N�V�[�g�����擾
    '// sheet_name = InputBox("�ǂݍ��ރ��[�N�V�[�g������͂��Ă��������B")
    target = "'" & open_file_name & sheet_name & "'!"
    
    ''���[�N�V�[�g�������������ǂ����A�܂��ǂݍ���ł݂�
    On Error Resume Next
    buf = ExecuteExcel4Macro(target & "R1C1")
    If Err <> 0 Then
        MsgBox "���[�N�V�[�g [ " & sheet_name & " ] ��ǂ߂܂���ł����B", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0 '// �G���[�����̖��ߎ�����

    Debug.Print target

End Sub

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�f�[�^�������Ă���J�n�s�E�I���s�����
'// ����    �F
'// �߂�l  �F
'// ���l    �F�J�����������݂���Ȃ�΁A�f�[�^�����J�n�s���牺���Ɍ�����i�߂čŏI�s����肷��B
'// �@�@�@�@�@�������ւ̌������A�󔒐����A�����Đݒ萔�����ł���Ό�����i�߂�B
'//----------------------------------------------------------------------------
Public Sub define_data_rows(ByVal field_name_ As String, field_row_ As Long, field_col_ As Long, data_first_row_ As Long, serch_end_blank_count_ As Long)

    Dim i As Long
    Dim buf As String
    Dim blank_count As Long

    '// �J�����������݂��Ȃ���ΏI��
    If is_existed_field_name(field_name_, field_row_, field_col_) = False Then
        MsgBox "Excel�䒠�ɃJ������ " & field_name_ & " �����݂��܂���B"
        Exit Sub
    End If

    ''�f�[�^�̓ǂݍ���
    i = 0
    blank_count = 0
    'buf = get_cell_value(data_first_row_ + i, field_col_)
    'If buf = "0" Then blank_count = blank_count + 1 Else blank_count = 0    '// �Z�����󔒂Ȃ�u�����N�J�E���g��+1
    
    Do While blank_count < serch_end_blank_count_
        
        '// �f�[�^��Ǎ�
        buf = ExecuteExcel4Macro(target & "R" & data_first_row_ + i & "C" & field_col_)
        
        If buf = "0" Then blank_count = blank_count + 1 Else blank_count = 0    '// �Z�����󔒂Ȃ�u�����N�J�E���g��+1
        
        i = i + 1
    Loop

    data_first_row = data_first_row_
    data_last_row = data_first_row_ + i - blank_count - 1
    If data_last_row < data_first_row Then data_last_row = data_first_row
    
    Debug.Print "�����ɂ��@�f�[�^�J�n�F" & data_first_row & "�@�f�[�^�I���F" & data_last_row

End Sub

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�w�肵��Exce�f�[�^��z��Ɋi�[���ĕԂ��֐�
'// ����    �F
'// �߂�l  �F
'// ���l    �F�@set_data_array �Apaste_data_to_excel ���Z�b�g�Ŏg���ꍇ�������B
'//----------------------------------------------------------------------------
Public Function set_data_array(ByVal field_name_ As String, field_row_ As Long, field_col_ As Long) As Variant()

    Dim i As Long
    Dim buf As String
     
    '// �J�����������݂��Ȃ���ΏI��
    If is_existed_field_name(field_name_, field_row_, field_col_) = False Then
        MsgBox "Excel�䒠�ɃJ������ " & field_name_ & " �����݂��܂���B"
        ReDim data(0)
        Exit Function
    End If

    '// �i�[�f�[�^���N���A
    Call Me.clear_data
    
    ''�f�[�^�̓ǂݍ���
    For i = data_first_row To data_last_row
    
        buf = get_cell_value(i, field_col_)
        Me.add_data buf
    
    Next
    
    set_data_array = data
    
End Function

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�w�肵��Excel�t�@�C���̓���Z���Ɋi�[���Ă����f�[�^���Z�b�g
'// ����    �F
'// �߂�l  �F
'// ���l    �F�@set_data_array �Apaste_data_to_excel ���Z�b�g�Ŏg���ꍇ�������B
'//----------------------------------------------------------------------------
Public Sub paste_data_to_excel(ByVal workbook_ As Workbook, sheet_name_ As String, row_ As Long, col_ As Long)

    Dim i As Long

    For i = 0 To Me.count_data
    
        workbook_.Worksheets(sheet_name_).Cells(row_ + i, col_) = Me.get_data_val(i)
        Debug.Print "i�F" & i & "�@�@�s�F" & row_ + i & "�@�@�l�F" & Me.get_data_val(i)
        
    Next
    Debug.Print ""

End Sub

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F��Excel�̃J�����������݂��邩���肷��֐�
'// ����    �F
'// �߂�l  �F
'// ���l    �FExcel�̍s�񂪕ς���Ă��Ȃ����ǂ����̃`�F�b�N�p
'//----------------------------------------------------------------------------
Private Function is_existed_field_name(ByVal field_name_ As String, row_ As Long, col_ As Long) As Boolean

    Dim buf As String
    
    buf = get_cell_value(row_, col_)
    
    If InStr(buf, field_name_) > 0 Then is_existed_field_name = True _
        Else is_existed_field_name = False

End Function

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F��Excel�̎w�肵���s��ɑ΂��āA�Z���l��Ԃ��֐�
'// ����    �F
'// �߂�l  �F
'// ���l    �F
'//----------------------------------------------------------------------------
Private Function get_cell_value(ByVal row_ As Long, ByVal col_ As Long) As String

    get_cell_value = ExecuteExcel4Macro(target & "R" & row_ & "C" & col_)

End Function

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�f�[�^�̒l��z��ɒǉ�
'// ����    �F
'// �߂�l  �F
'// ���l    �F
'//----------------------------------------------------------------------------
Public Sub add_data(ByVal val)
    On Error Resume Next
    
    Dim i
    
    i = Me.count_data
    
    If (IsEmpty(data(i)) = True) Then
        data(i) = val
    Else
        ReDim Preserve data(i + 1)
        data(i + 1) = val
    End If
 
End Sub

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�w��v�f�ʒu�̒l���擾����
'// ����    �F
'// �߂�l  �F�w��v�f�ʒu�̒l
'// ���l    �FJava��ArrayList.Get�ƈꏏ
'//----------------------------------------------------------------------------
Public Function get_data_val(index)
    Dim v_ret As Variant
    
    If (index > Me.count_data) Then
        v_ret = Null
    Else
        v_ret = data(index)
    End If
    
    get_data_val = v_ret
End Function

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�f�[�^�z��̗v�f���𒲂ׂ�B
'// ����    �F
'// �߂�l  �F
'// ���l    �F
'//----------------------------------------------------------------------------
Public Function count_data() As Long
    count_data = UBound(data)
End Function

'//----------------------------------------------------------------------------
'// �֐���  �F
'// �@�\    �F�f�[�^�z����N���A����B
'// ����    �F
'// �߂�l  �F
'// ���l    �F
'//----------------------------------------------------------------------------
Public Sub clear_data()
    ReDim data(0)
End Sub