Attribute VB_Name = "Module1"
Option Explicit

Const MAIN_SHEET = "Main"       '// データ格納がされるシート（今回はMainシート）
Const CONFIG_SHEET = "Config"

Const MAIN_SQL_TYPE_DEFINED_ROW = 1         '// SQLタイプを定義しているMainシートの行
Const MAIN_DB_COL_NAME_DEFINED_ROW = 2      '// DBの列名を定義しているMainシートの行

Const MAIN_DATA_FIELD_NAME_DEFINED_ROW = 3          '// 抽出するExcelデータのフィールド名に含まれる文字列を定義しているMainシートの行
Const MAIN_DATA_FIELD_ROW_DEFINED_ROW = 4           '// 抽出するExcelデータのフィールド名が含まれる行を定義しているMainシートの行
Const MAIN_DATA_FIELD_COL_DEFINED_ROW = 5           '// 抽出するExcelデータの列を定義しているMainシートの行
Const MAIN_DATA_FORMAT_DEFINED_ROW = 6              '// 抽出するExcelデータの書式を定義しているMainシートの行

Const MAIN_DATA_START_ROW = 7       '// Mainシートのデータ格納開始行
Const MAIN_DATA_START_COL = 10      '// Mainシートのデータ格納開始列

Dim excel_data_start_row As Long    '// Excel台帳のデータ開始行
Dim excel_data_end_row As Long      '// Excel台帳のデータ終了行（検索により決定）

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

Dim common As ExcelCommon                       '// Excelで良く使う関数が含まれるクラスをインスタンス化
Dim excel_data_getter As ExcelDataGetter
Dim sql_list As ArrayList
Dim ado As AdodbInterface

Dim sql_mode As Long    '//1:エラー終了　2:INSERT　3:UPDATE

'//----------------------------------------------------------------------------
'// 機能    ：メイン実行プログラム
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub main()
    
    '// オブジェクトをインスタンス化
    Set common = New ExcelCommon
    Set excel_data_getter = New ExcelDataGetter
    Set sql_list = New ArrayList
    Set ado = New AdodbInterface
    
    '// ExcelCommonにこのワークブックをセット
    common.set_workbook ThisWorkbook
    
    '// Excel台帳やDBの設定情報取得
    Call set_info
    
    '// SQL区分設定のチェック
    sql_mode = check_sqltype
    If sql_mode = 1 Then End  '// 戻り値が1なら終了
    
    '// Excel台帳データ用のセルクリア
    Call common.data_clear( _
        MAIN_SHEET _
        , MAIN_DATA_START_ROW _
        , MAIN_DATA_START_COL _
        , common.get_max_row_below(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL) _
        , common.get_max_col_right(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL))
    
    
    '// データ取得するExcel台帳を設定
    excel_data_getter.set_info EXCEL_FOLDER, EXCEL_FILE, EXCEL_SHEET
    
    '// Excel台帳をオープン
    common.open_file EXCEL_FOLDER & "\" & EXCEL_FILE
    
    '// Excel台帳からデータ取得する行の開始と終了を決定
    excel_data_getter.define_data_rows EXCEL_SEARCH_FIELD_NAME, EXCEL_SEARCH_FIELD_NAME_ROW, EXCEL_SEARCH_FIELD_NAME_COL, _
        EXCEL_SEARCH_DATA_START_ROW, EXCEL_SEARCH_END_MULTI_BLANK
    
    '// Excel台帳からデータ取得
    paste_excel_data MAIN_SHEET
    
    '// Excel台帳をクローズ
    common.close_opened_file False
    
    '// 取得したデータからSQLを生成して実行
    '//Call make_sql_list(sql_list, MAIN_SHEET)
    '//Call excute_sql(sql_list)
    Call make_and_excute_sqls
    
    '// DBへの切断処理
    ado.close_connection
    
    MsgBox "Excel台帳データをDBへ反映させました。"

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：Excelから基本情報をセット
'// 備考    ：
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
'// 機能    ：ExcelのSQLタイプ設定が正しいかチェックする。
'// 備考    ：
'//----------------------------------------------------------------------------
Private Function check_sqltype() As Long

    Dim count_insert As Long, count_update As Long, count_where As Long
    
    Dim sqltype As String
    Dim row_end As Long
    
    Dim i As Long
    
    check_sqltype = 0   '// 問題がある場合は1の戻り値を返す
    
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
    
    
    '// INSERT文を扱う場合
    If count_insert > 0 And count_update = 0 And count_where = 0 Then
        check_sqltype = 2   '// SQLタイプ設定はOK:INSERTモード
    
    '// UPDATE文を扱う場合
    ElseIf count_insert = 0 And count_update > 0 And count_where > 0 Then
        check_sqltype = 3   '// SQLタイプ設定はOK:UPDATEモード
    
    '// 上記以外は中断処理
    Else
        check_sqltype = 1
        MsgBox ("SQLタイプエラー：処理を中断します。" & Chr(13) _
            & "　INSERT文を扱う場合は、UPDATEとWHEREは指定しないで下さい。" & Chr(13) _
            & "　UPDATE文を扱う場合は、INSERTを指定せずにWHEREを1つ以上指定して下さい。")
    
    End If

End Function

'//----------------------------------------------------------------------------
'// 機能    ：Excel台帳からデータをセットする。
'// 備考    ：フォルダ名、ファイル名、シート名を引数としておくことで、複数ファイルからの同時取り込みへと拡張可能
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
    
        '// Excel台帳からデータを取得するための条件取得
        field_name = ThisWorkbook.Worksheets(sheet_name_).Cells(MAIN_DATA_FIELD_NAME_DEFINED_ROW, MAIN_DATA_START_COL + i)
        field_row = ThisWorkbook.Worksheets(sheet_name_).Cells(MAIN_DATA_FIELD_ROW_DEFINED_ROW, MAIN_DATA_START_COL + i)
        field_col = ThisWorkbook.Worksheets(sheet_name_).Cells(MAIN_DATA_FIELD_COL_DEFINED_ROW, MAIN_DATA_START_COL + i)
        
        '// Excel台帳のデータを貼り付け
        excel_data_getter.set_data_array field_name, field_row, field_col
        excel_data_getter.paste_data_to_excel ThisWorkbook, sheet_name_, MAIN_DATA_START_ROW, MAIN_DATA_START_COL + i
        
        '// Excel台帳から取得したデータの書式を変更
        format = ThisWorkbook.Worksheets(sheet_name_).Cells(MAIN_DATA_FORMAT_DEFINED_ROW, MAIN_DATA_START_COL + i)
        common.change_format_max_row_below sheet_name_, MAIN_DATA_START_ROW, MAIN_DATA_START_COL + i, format
    
        '// 次の処理へ移る準備
        i = i + 1
        sqltype = ThisWorkbook.Worksheets(sheet_name_).Cells(1, MAIN_DATA_START_COL + i).Value
    Loop

End Function

'//----------------------------------------------------------------------------
'// 機能    ：格納されたデータからSQL文を作成しリストに追加する。
'// 備考    ：
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
    
    Debug.Print ("リスト内保有数： " & list.count() + 1)
    
    For i = 0 To list.count()
        '// Debug.Print i & ":" & list.GetVal(i)
    Next
    
End Sub

'//----------------------------------------------------------------------------
'// 機能    ：リスト内のSQL文を実行する。
'// 備考    ：
'//----------------------------------------------------------------------------
Private Sub excute_sql(ByVal sql_list As ArrayList)

    Dim i As Long
        
    '// Debug.Print (DB_DRIVER & ", "&DB_NETSERVICENAME & ", " & DB_DSN & ", " & DB_USER & ", " & DB_PASSWORD)
    ado.open_oracle DB_DRIVER, DB_NETSERVICENAME, DB_DSN, DB_USER, DB_PASSWORD, DB_CONNECT_MODE    '// ()で引数渡しをするとエラー発生

    ado.con_begintrans
    For i = 0 To sql_list.count
        '// Debug.Print sql_list.GetVal(i)
        ado.excute_sql sql_list.GetVal(i)
    Next
    ado.con_committrans

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：格納されたデータからSQL文を作成し実行する。
'// 備考    ：
'//----------------------------------------------------------------------------
Private Sub make_and_excute_sqls()

    Dim end_col As Long
    end_col = common.get_max_col_right(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL)

    ado.open_oracle DB_DRIVER, DB_NETSERVICENAME, DB_DSN, DB_USER, DB_PASSWORD, DB_CONNECT_MODE    '// ()で引数渡しをするとエラー発生
    ado.make_and_excute_sqls DB_TABLE, MAIN_SHEET, MAIN_SQL_TYPE_DEFINED_ROW, MAIN_DB_COL_NAME_DEFINED_ROW, MAIN_DATA_START_ROW, MAIN_DATA_START_COL, end_col

End Sub


'//----------------------------------------------------------------------------
'// 以降、デバッグ用関数
'//----------------------------------------------------------------------------

'//----------------------------------------------------------------------------
'// 機能    ：オラクルDBへの接続テスト
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub debug_connect_oracle()

    '// Excel台帳やDBの設定情報取得
    Call set_info
    
    Set ado = New AdodbInterface
    ado.open_oracle DB_DRIVER, DB_NETSERVICENAME, DB_DSN, DB_USER, DB_PASSWORD, DB_CONNECT_MODE   '// ()で引数渡しをするとエラー発生
    ado.close_connection
    
    MsgBox "オラクルへの接続テスト完了！"

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：Excel台帳からデータをセットするまでのテスト
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub debug_paste_excel_data()

    '// オブジェクトをインスタンス化
    Set common = New ExcelCommon
    Set excel_data_getter = New ExcelDataGetter
    Set sql_list = New ArrayList
    Set ado = New AdodbInterface
    
    '// ExcelCommonにこのワークブックをセット
    common.set_workbook ThisWorkbook
    
    '// Excel台帳やDBの設定情報取得
    Call set_info
    
    '// SQL区分設定のチェック
    sql_mode = check_sqltype
    If sql_mode = 1 Then End  '// 戻り値が1なら終了
    
    '// Excel台帳データ用のセルクリア
    Call common.data_clear( _
        MAIN_SHEET _
        , MAIN_DATA_START_ROW _
        , MAIN_DATA_START_COL _
        , common.get_max_row_below(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL) _
        , common.get_max_col_right(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL))
    
    
    '// データ取得するExcel台帳を設定
    excel_data_getter.set_info EXCEL_FOLDER, EXCEL_FILE, EXCEL_SHEET
    
    '// Excel台帳をオープン
    common.open_file EXCEL_FOLDER & "\" & EXCEL_FILE
    
    '// Excel台帳からデータ取得する行の開始と終了を決定
    excel_data_getter.define_data_rows EXCEL_SEARCH_FIELD_NAME, EXCEL_SEARCH_FIELD_NAME_ROW, EXCEL_SEARCH_FIELD_NAME_COL, _
        EXCEL_SEARCH_DATA_START_ROW, EXCEL_SEARCH_END_MULTI_BLANK
    
    '// Excel台帳からデータ取得
    paste_excel_data MAIN_SHEET
    
    '// Excel台帳をクローズ
    common.close_opened_file False
    
End Sub

'//----------------------------------------------------------------------------
'// 機能    ：SQLを作成して実行するテスト
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub debug_excute_sql()

    Dim rc As Integer
    rc = MsgBox("処理を行いますか？", vbYesNo + vbQuestion, "確認")
    If rc = vbYes Then
        '//MsgBox "処理を行います"
    Else
        MsgBox "処理を中断します"
        Exit Sub
    End If

    '// オブジェクトをインスタンス化
    Set common = New ExcelCommon
    Set excel_data_getter = New ExcelDataGetter
    Set sql_list = New ArrayList
    Set ado = New AdodbInterface
    
    '// ExcelCommonにこのワークブックをセット
    common.set_workbook ThisWorkbook
    
    '// Excel台帳やDBの設定情報取得
    Call set_info
    
    '// SQL区分設定のチェック
    sql_mode = check_sqltype
    If sql_mode = 1 Then End  '// 戻り値が1なら終了
    
    
    '// 取得したデータからSQLを生成して実行
    Call make_and_excute_sqls
    
    MsgBox "Excel台帳データをDBへ反映させました。"


End Sub

'//----------------------------------------------------------------------------
'// 機能    ：格納しているデータをクリア
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub debug_data_clear()

    '// オブジェクトをインスタンス化
    Set common = New ExcelCommon

    '// ExcelCommonにこのワークブックをセット
    common.set_workbook ThisWorkbook
    
    '// Excel台帳やDBの設定情報取得
    Call set_info
        
    '// Excel台帳データ用のセルクリア
    Call common.data_clear( _
        MAIN_SHEET _
        , MAIN_DATA_START_ROW _
        , MAIN_DATA_START_COL _
        , common.get_max_row_below(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL) _
        , common.get_max_col_right(MAIN_SHEET, MAIN_DATA_START_ROW, MAIN_DATA_START_COL))
        
End Sub

'//----------------------------------------------------------------------------
'// 機能    ：テスト用の一時的な関数
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub temp()

    '// オブジェクトをインスタンス化
    Set common = New ExcelCommon
    Set excel_data_getter = New ExcelDataGetter
    Set sql_list = New ArrayList
    Set ado = New AdodbInterface
    
    '// ExcelCommonにこのワークブックをセット
    common.set_workbook ThisWorkbook
    
    '// Excel台帳やDBの設定情報取得
    Call set_info

    Debug.Print common.exist_val(9, "Sheet1", 1, 1)

End Sub

