Attribute VB_Name = "Module1"
Option Explicit

Const MAIN_SHEET = "Main"       '// �f�[�^�i�[�������V�[�g�i�����Main�V�[�g�j
Const CONFIG_SHEET = "Config"

Const MAIN_SQL_TYPE_DEFINED_ROW = 1         '// SQL�^�C�v���`���Ă���Main�V�[�g�̍s
Const MAIN_DB_COL_NAME_DEFINED_ROW = 2      '// DB�̗񖼂��`���Ă���Main�V�[�g�̍s

Const MAIN_DATA_FIELD_NAME_DEFINED_ROW = 3          '// ���o����Excel�f�[�^�̃t�B�[���h���Ɋ܂܂�镶������`���Ă���Main�V�[�g�̍s
Const MAIN_DATA_FIELD_ROW_DEFINED_ROW = 4           '// ���o����Excel�f�[�^�̃t�B�[���h�����܂܂��s���`���Ă���Main�V�[�g�̍s
Const MAIN_DATA_FIELD_COL_DEFINED_ROW = 5           '// ���o����Excel�f�[�^�̗���`���Ă���Main�V�[�g�̍s
Const MAIN_DATA_FORMAT_DEFINED_ROW = 6              '// ���o����Excel�f�[�^�̏������`���Ă���Main�V�[�g�̍s

Const MAIN_DATA_START_ROW = 7       '// Main�V�[�g�̃f�[�^�i�[�J�n�s
Const MAIN_DATA_START_COL = 10      '// Main�V�[�g�̃f�[�^�i�[�J�n��

Dim excel_data_start_row As Long    '// Excel�䒠�̃f�[�^�J�n�s
Dim excel_data_end_row As Long      '// Excel�䒠�̃f�[�^�I���s�i�����ɂ�茈��j

Dim EXCEL_FOLDER As String
Dim EXCEL_FILE As String
Dim EXCEL_SHEET As String

Dim EXCEL_SEARCH_FIELD_NAME As String
Dim EXCEL_SEARCH_FIELD_NAME_ROW As Long
Dim EXCEL_SEARCH_FIELD_NAME_COL As Long
Dim EXCEL_SEARCH_DATA_START_ROW As Long
Dim EXCEL_SEARCH_END_MULTI_BLANK As Long

Dim DB_CONNECT_MODE As Long
Dim DB_DRIVER As String
Dim DB_NETSERVICENAME As String
Dim DB_DSN As String
Dim DB_USER As String
Dim DB_PASSWORD As String
Dim DB_TABLE As String

Dim common As ExcelCommon                       '// Excel�ŗǂ��g���֐����܂܂��N���X���C���X�^���X��
Dim excel_data_getter As ExcelDataGetter
Dim sql_list As ArrayList
Dim ado As AdodbInterface

Dim sql_mode As Long    '//1:�G���[�I���@2:INSERT�@3:UPDATE

'//----------------------------------------------------------------------------
'// �@�\    �F���C�����s�v���O����
'// ���l    �F
'//----------------------------------------------------------------------------
Public Sub main()
    
    '// �I�u�W�F�N�g���C���X�^���X��
    Set common = New ExcelCommon
    Set excel_data_getter = New ExcelDataGetter
    Set sql_list = New ArrayList
    Set ado = New AdodbInterface
    
    '// ExcelCommon�ɂ��̃��[�N�u�b�N���Z�b�g
    common.set_workbook ThisWorkbook
    
    '// Excel�䒠��DB�̐ݒ���擾
    Call set_info
    
    '// SQL�敪�ݒ�̃`�F�b�N
    sql_mode = check_sqltype
    If sql_mode = 1 Then End  '// �߂�l��1�Ȃ�I��
    
    '// Excel�䒠�f�[�^�p�̃Z���N���A
    Call common.data_clear( _
        MAIN_SHEET _
        , MAIN_DATA_START_ROW _
        , MAIN_DATA_START_COL _
        , common.get_max_row_below(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL) _
        , common.get_max_col_right(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL))
    
    
    '// �f�[�^�擾����Excel�䒠��ݒ�
    excel_data_getter.set_info EXCEL_FOLDER, EXCEL_FILE, EXCEL_SHEET
    
    '// Excel�䒠���I�[�v��
    common.open_file EXCEL_FOLDER & "\" & EXCEL_FILE
    
    '// Excel�䒠����f�[�^�擾����s�̊J�n�ƏI��������
    excel_data_getter.define_data_rows EXCEL_SEARCH_FIELD_NAME, EXCEL_SEARCH_FIELD_NAME_ROW, EXCEL_SEARCH_FIELD_NAME_COL, _
        EXCEL_SEARCH_DATA_START_ROW, EXCEL_SEARCH_END_MULTI_BLANK
    
    '// Excel�䒠����f�[�^�擾
    paste_excel_data MAIN_SHEET
    
    '// Excel�䒠���N���[�Y
    common.close_opened_file False
    
    '// �擾�����f�[�^����SQL�𐶐����Ď��s
    '//Call make_sql_list(sql_list, MAIN_SHEET)
    '//Call excute_sql(sql_list)
    Call make_and_excute_sqls
    
    '// DB�ւ̐ؒf����
    ado.close_connection
    
    MsgBox "Excel�䒠�f�[�^��DB�֔��f�����܂����B"

End Sub

'//----------------------------------------------------------------------------
'// �@�\    �FExcel�����{�����Z�b�g
'// ���l    �F
'//----------------------------------------------------------------------------
Private Sub set_info()

    EXCEL_FOLDER = ThisWorkbook.Worksheets(CONFIG_SHEET).Cells(2, 4).Value
    EXCEL_FILE = ThisWorkbook.Worksheets(CONFIG_SHEET).Cells(3, 4).Value
    EXCEL_SHEET = ThisWorkbook.Worksheets(CONFIG_SHEET).Cells(4, 4).Value
    
    EXCEL_SEARCH_FIELD_NAME = ThisWorkbook.Worksheets(CONFIG_SHEET).Cells(7, 4).Value
    EXCEL_SEARCH_FIELD_NAME_ROW = ThisWorkbook.Worksheets(CONFIG_SHEET).Cells(8, 4).Value
    EXCEL_SEARCH_FIELD_NAME_COL = ThisWorkbook.Worksheets(CONFIG_SHEET).Cells(9, 4).Value
    EXCEL_SEARCH_DATA_START_ROW = ThisWorkbook.Worksheets(CONFIG_SHEET).Cells(10, 4).Value
    EXCEL_SEARCH_END_MULTI_BLANK = ThisWorkbook.Worksheets(CONFIG_SHEET).Cells(11, 4).Value
    
    DB_CONNECT_MODE = ThisWorkbook.Worksheets(CONFIG_SHEET).Cells(2, 7).Value
    DB_DRIVER = ThisWorkbook.Worksheets(CONFIG_SHEET).Cells(3, 7).Value
    DB_NETSERVICENAME = ThisWorkbook.Worksheets(CONFIG_SHEET).Cells(4, 7).Value
    DB_DSN = ThisWorkbook.Worksheets(CONFIG_SHEET).Cells(5, 7).Value
    DB_USER = ThisWorkbook.Worksheets(CONFIG_SHEET).Cells(6, 7).Value
    DB_PASSWORD = ThisWorkbook.Worksheets(CONFIG_SHEET).Cells(7, 7).Value
    DB_TABLE = ThisWorkbook.Worksheets(CONFIG_SHEET).Cells(8, 7).Value
        
End Sub

'//----------------------------------------------------------------------------
'// �@�\    �FExcel��SQL�^�C�v�ݒ肪���������`�F�b�N����B
'// ���l    �F
'//----------------------------------------------------------------------------
Private Function check_sqltype() As Long

    Dim count_insert As Long, count_update As Long, count_where As Long
    
    Dim sqltype As String
    Dim row_end As Long
    
    Dim i As Long
    
    check_sqltype = 0   '// ��肪����ꍇ��1�̖߂�l��Ԃ�
    
    count_insert = 0
    count_update = 0
    count_where = 0
    row_end = common.get_max_col_right(MAIN_SHEET, 1, MAIN_DATA_START_COL)
    
    For i = MAIN_DATA_START_COL To row_end
        Select Case ThisWorkbook.Worksheets(MAIN_SHEET).Cells(1, i)
        
            Case "INSERT"
                count_insert = count_insert + 1
    
            Case "UPDATE"
                count_update = count_update + 1
        
            Case "WHERE"
                count_where = count_where + 1
        End Select
    Next
    
    Debug.Print (count_insert & ", " & count_update & ", " & count_where)
    
    
    '// INSERT���������ꍇ
    If count_insert > 0 And count_update = 0 And count_where = 0 Then
        check_sqltype = 2   '// SQL�^�C�v�ݒ��OK:INSERT���[�h
    
    '// UPDATE���������ꍇ
    ElseIf count_insert = 0 And count_update > 0 And count_where > 0 Then
        check_sqltype = 3   '// SQL�^�C�v�ݒ��OK:UPDATE���[�h
    
    '// ��L�ȊO�͒��f����
    Else
        check_sqltype = 1
        MsgBox ("SQL�^�C�v�G���[�F�����𒆒f���܂��B" & Chr(13) _
            & "�@INSERT���������ꍇ�́AUPDATE��WHERE�͎w�肵�Ȃ��ŉ������B" & Chr(13) _
            & "�@UPDATE���������ꍇ�́AINSERT���w�肹����WHERE��1�ȏ�w�肵�ĉ������B")
    
    End If

End Function

'//----------------------------------------------------------------------------
'// �@�\    �FExcel�䒠����f�[�^���Z�b�g����B
'// ���l    �F�t�H���_���A�t�@�C�����A�V�[�g���������Ƃ��Ă������ƂŁA�����t�@�C������̓�����荞�݂ւƊg���\
'//----------------------------------------------------------------------------
Private Function paste_excel_data(sheet_name_ As String)
     
    Dim i As Long
    Dim sqltype As String

    Dim field_name As String
    Dim field_row As Long
    Dim field_col As Long
    Dim format As String

    i = 0
    sqltype = ThisWorkbook.Worksheets(sheet_name_).Cells(1, MAIN_DATA_START_COL).Value
    Do While sqltype <> ""
    
        '// Excel�䒠����f�[�^���擾���邽�߂̏����擾
        field_name = ThisWorkbook.Worksheets(sheet_name_).Cells(MAIN_DATA_FIELD_NAME_DEFINED_ROW, MAIN_DATA_START_COL + i)
        field_row = ThisWorkbook.Worksheets(sheet_name_).Cells(MAIN_DATA_FIELD_ROW_DEFINED_ROW, MAIN_DATA_START_COL + i)
        field_col = ThisWorkbook.Worksheets(sheet_name_).Cells(MAIN_DATA_FIELD_COL_DEFINED_ROW, MAIN_DATA_START_COL + i)
        
        '// Excel�䒠�̃f�[�^��\��t��
        excel_data_getter.set_data_array field_name, field_row, field_col
        excel_data_getter.paste_data_to_excel ThisWorkbook, sheet_name_, MAIN_DATA_START_ROW, MAIN_DATA_START_COL + i
        
        '// Excel�䒠����擾�����f�[�^�̏�����ύX
        format = ThisWorkbook.Worksheets(sheet_name_).Cells(MAIN_DATA_FORMAT_DEFINED_ROW, MAIN_DATA_START_COL + i)
        common.change_format_max_row_below sheet_name_, MAIN_DATA_START_ROW, MAIN_DATA_START_COL + i, format
    
        '// ���̏����ֈڂ鏀��
        i = i + 1
        sqltype = ThisWorkbook.Worksheets(sheet_name_).Cells(1, MAIN_DATA_START_COL + i).Value
    Loop

End Function

'//----------------------------------------------------------------------------
'// �@�\    �F�i�[���ꂽ�f�[�^����SQL�����쐬�����X�g�ɒǉ�����B
'// ���l    �F
'//----------------------------------------------------------------------------
Private Sub make_sql_list(ByVal list As ArrayList, sheet_name As String)

    Dim i As Long, j As Long
    Dim start_row As Long, end_row As Long
    Dim start_col As Long, end_col As Long
    
    Dim sqltype As String
    Dim sql As String
    
    start_row = MAIN_DATA_START_ROW
    end_row = common.get_max_row_below(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL)

    start_col = MAIN_DATA_START_COL
    end_col = common.get_max_col_right(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL)
    
             
    For i = start_row To end_row
    
        sql = ado.make_sql(DB_TABLE, sheet_name, MAIN_SQL_TYPE_DEFINED_ROW, MAIN_DB_COL_NAME_DEFINED_ROW, i, start_col, end_col)
        list.add (sql)
        
    Next
    
    Debug.Print ("���X�g���ۗL���F " & list.count() + 1)
    
    For i = 0 To list.count()
        '// Debug.Print i & ":" & list.GetVal(i)
    Next
    
End Sub

'//----------------------------------------------------------------------------
'// �@�\    �F���X�g����SQL�������s����B
'// ���l    �F
'//----------------------------------------------------------------------------
Private Sub excute_sql(ByVal sql_list As ArrayList)

    Dim i As Long
        
    '// Debug.Print (DB_DRIVER & ", "&DB_NETSERVICENAME & ", " & DB_DSN & ", " & DB_USER & ", " & DB_PASSWORD)
    ado.open_oracle DB_DRIVER, DB_NETSERVICENAME, DB_DSN, DB_USER, DB_PASSWORD, DB_CONNECT_MODE    '// ()�ň����n��������ƃG���[����

    ado.con_begintrans
    For i = 0 To sql_list.count
        '// Debug.Print sql_list.GetVal(i)
        ado.excute_sql sql_list.GetVal(i)
    Next
    ado.con_committrans

End Sub

'//----------------------------------------------------------------------------
'// �@�\    �F�i�[���ꂽ�f�[�^����SQL�����쐬�����s����B
'// ���l    �F
'//----------------------------------------------------------------------------
Private Sub make_and_excute_sqls()

    Dim end_col As Long
    end_col = common.get_max_col_right(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL)

    ado.open_oracle DB_DRIVER, DB_NETSERVICENAME, DB_DSN, DB_USER, DB_PASSWORD, DB_CONNECT_MODE    '// ()�ň����n��������ƃG���[����
    ado.make_and_excute_sqls DB_TABLE, MAIN_SHEET, MAIN_SQL_TYPE_DEFINED_ROW, MAIN_DB_COL_NAME_DEFINED_ROW, MAIN_DATA_START_ROW, MAIN_DATA_START_COL, end_col

End Sub


'//----------------------------------------------------------------------------
'// �ȍ~�A�f�o�b�O�p�֐�
'//----------------------------------------------------------------------------

'//----------------------------------------------------------------------------
'// �@�\    �F�I���N��DB�ւ̐ڑ��e�X�g
'// ���l    �F
'//----------------------------------------------------------------------------
Public Sub debug_connect_oracle()

    '// Excel�䒠��DB�̐ݒ���擾
    Call set_info
    
    Set ado = New AdodbInterface
    ado.open_oracle DB_DRIVER, DB_NETSERVICENAME, DB_DSN, DB_USER, DB_PASSWORD, DB_CONNECT_MODE   '// ()�ň����n��������ƃG���[����
    ado.close_connection
    
    MsgBox "�I���N���ւ̐ڑ��e�X�g�����I"

End Sub

'//----------------------------------------------------------------------------
'// �@�\    �FExcel�䒠����f�[�^���Z�b�g����܂ł̃e�X�g
'// ���l    �F
'//----------------------------------------------------------------------------
Public Sub debug_paste_excel_data()

    '// �I�u�W�F�N�g���C���X�^���X��
    Set common = New ExcelCommon
    Set excel_data_getter = New ExcelDataGetter
    Set sql_list = New ArrayList
    Set ado = New AdodbInterface
    
    '// ExcelCommon�ɂ��̃��[�N�u�b�N���Z�b�g
    common.set_workbook ThisWorkbook
    
    '// Excel�䒠��DB�̐ݒ���擾
    Call set_info
    
    '// SQL�敪�ݒ�̃`�F�b�N
    sql_mode = check_sqltype
    If sql_mode = 1 Then End  '// �߂�l��1�Ȃ�I��
    
    '// Excel�䒠�f�[�^�p�̃Z���N���A
    Call common.data_clear( _
        MAIN_SHEET _
        , MAIN_DATA_START_ROW _
        , MAIN_DATA_START_COL _
        , common.get_max_row_below(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL) _
        , common.get_max_col_right(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL))
    
    
    '// �f�[�^�擾����Excel�䒠��ݒ�
    excel_data_getter.set_info EXCEL_FOLDER, EXCEL_FILE, EXCEL_SHEET
    
    '// Excel�䒠���I�[�v��
    common.open_file EXCEL_FOLDER & "\" & EXCEL_FILE
    
    '// Excel�䒠����f�[�^�擾����s�̊J�n�ƏI��������
    excel_data_getter.define_data_rows EXCEL_SEARCH_FIELD_NAME, EXCEL_SEARCH_FIELD_NAME_ROW, EXCEL_SEARCH_FIELD_NAME_COL, _
        EXCEL_SEARCH_DATA_START_ROW, EXCEL_SEARCH_END_MULTI_BLANK
    
    '// Excel�䒠����f�[�^�擾
    paste_excel_data MAIN_SHEET
    
    '// Excel�䒠���N���[�Y
    common.close_opened_file False
    
End Sub

'//----------------------------------------------------------------------------
'// �@�\    �FSQL���쐬���Ď��s����e�X�g
'// ���l    �F
'//----------------------------------------------------------------------------
Public Sub debug_excute_sql()

    Dim rc As Integer
    rc = MsgBox("�������s���܂����H", vbYesNo + vbQuestion, "�m�F")
    If rc = vbYes Then
        '//MsgBox "�������s���܂�"
    Else
        MsgBox "�����𒆒f���܂�"
        Exit Sub
    End If

    '// �I�u�W�F�N�g���C���X�^���X��
    Set common = New ExcelCommon
    Set excel_data_getter = New ExcelDataGetter
    Set sql_list = New ArrayList
    Set ado = New AdodbInterface
    
    '// ExcelCommon�ɂ��̃��[�N�u�b�N���Z�b�g
    common.set_workbook ThisWorkbook
    
    '// Excel�䒠��DB�̐ݒ���擾
    Call set_info
    
    '// SQL�敪�ݒ�̃`�F�b�N
    sql_mode = check_sqltype
    If sql_mode = 1 Then End  '// �߂�l��1�Ȃ�I��
    
    
    '// �擾�����f�[�^����SQL�𐶐����Ď��s
    Call make_and_excute_sqls
    
    MsgBox "Excel�䒠�f�[�^��DB�֔��f�����܂����B"


End Sub

'//----------------------------------------------------------------------------
'// �@�\    �F�i�[���Ă���f�[�^���N���A
'// ���l    �F
'//----------------------------------------------------------------------------
Public Sub debug_data_clear()

    '// �I�u�W�F�N�g���C���X�^���X��
    Set common = New ExcelCommon

    '// ExcelCommon�ɂ��̃��[�N�u�b�N���Z�b�g
    common.set_workbook ThisWorkbook
    
    '// Excel�䒠��DB�̐ݒ���擾
    Call set_info
        
    '// Excel�䒠�f�[�^�p�̃Z���N���A
    Call common.data_clear( _
        MAIN_SHEET _
        , MAIN_DATA_START_ROW _
        , MAIN_DATA_START_COL _
        , common.get_max_row_below(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL) _
        , common.get_max_col_right(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL))
        
End Sub

'//----------------------------------------------------------------------------
'// �@�\    �F�e�X�g�p�̈ꎞ�I�Ȋ֐�
'// ���l    �F
'//----------------------------------------------------------------------------
Public Sub temp()

    '// �I�u�W�F�N�g���C���X�^���X��
    Set common = New ExcelCommon
    Set excel_data_getter = New ExcelDataGetter
    Set sql_list = New ArrayList
    Set ado = New AdodbInterface
    
    '// ExcelCommon�ɂ��̃��[�N�u�b�N���Z�b�g
    common.set_workbook ThisWorkbook
    
    '// Excel�䒠��DB�̐ݒ���擾
    Call set_info

    Debug.Print common.exist_val(9, "Sheet1", 1, 1)

End Sub

