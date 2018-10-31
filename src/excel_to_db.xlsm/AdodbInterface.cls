VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdodbInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//----------------------------------------------------------------------------
'// AdodbInterfaceクラス
'//----------------------------------------------------------------------------

Option Explicit

'//----------------------------------------------------------------------------
'// 【使用例】
'//  ①インスタンス生成
'//　②open_oracle：オラクルへの接続 or open_access：アクセスへの接続
'//　③excute_sql or excute_sql_and_paste_recordset：SQLの実行やデータ貼り付け
'//　④paste_recordset：実行したSQLにより取得したデータを貼り付け
'//　⑤close_connection：DBへの切断
'//----------------------------------------------------------------------------

'// 定義エラー発生時は参照設定で[Microsoft ActiveX Data Objects x.x Library]が選択されているか確認
Private con As Adodb.Connection
Private rs As Adodb.Recordset

Private driver As String
Private netservicename As String
Private dsn As String
Private username As String
Private password As String

'//----------------------------------------------------------------------------
'// 機能    ：コンストラクタ
'// 備考    ：
'//----------------------------------------------------------------------------
Private Sub Class_Initialize()
    
    Set con = New Adodb.Connection
    Set rs = New Adodb.Recordset
    
    
End Sub

'//----------------------------------------------------------------------------
'// 機能    ：オラクルへの接続処理
'// 備考    ：接続モード 1 ： ドライバー名・ネットサービス名を指
'//　　　　　 　　　　　 2 ： データソース名を指定
'//----------------------------------------------------------------------------
Public Sub open_oracle(ByVal driver_ As String, netservicename_ As String, dsn_ As String, username_ As String, password_ As String, connect_mode_ As Long)
    
    Dim constr As String

    driver = "{" & driver_ & "}"        'ex.ドライバ名（Microsoft社製）→{Microsoft ODBC for Oracle}, ドライバ名（Oracle社製）→{Oracle in OraDb11g_home1}
    netservicename = netservicename_    'tnsnames.ora ファイルのネットサービス名'
    dsn = dsn_                          'データソース名
    username = username_                '接続するデータベースのユーザー名
    password = password_                'パスワード

    'ドライバー名・ネットサービス名を指定する場合
    If connect_mode_ = 1 Then
        constr = "DRIVER=" & driver
        constr = constr & ";CONNECTSTRING=" & netservicename
        constr = constr & ";UID=" & username
        constr = constr & ";PWD=" & password & ";"
    ElseIf connect_mode_ = 2 Then
        'データソース名を指定する場合
        constr = "DSN=" & dsn
        constr = constr & ";UID=" & username
        constr = constr & ";PWD=" & password
    Else
        MsgBox "用意されていないDB接続モードが指定されています。" & Chr(13) & "設定を確認して下さい。"
    End If
    
    Debug.Print (constr)
    con.ConnectionString = constr
    con.Open
    Debug.Print "オラクルへの接続完了"
    
End Sub

'//----------------------------------------------------------------------------
'// 機能    ：アクセスへの接続処理
'// 備考    ：引数 access_name_ はフォルダ名＋ファイル名
'//----------------------------------------------------------------------------
Public Sub open_access(ByVal access_name_ As String)
     
    con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & access_name_ & ";" 'Accessファイルに接続
    Debug.Print "アクセスへの接続完了"
     
End Sub

'//----------------------------------------------------------------------------
'// 機能    ：SQLの実行
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub excute_sql(ByVal str_sql_ As String)
    
    On Error GoTo err_
    
    '// UPDATE文を使用する際は、事前にCOMMITさせておかないとエラーになるため注意
    Debug.Print str_sql_ & " を実行します。"
    rs.Open str_sql_, con
    
    Exit Sub

err_:
    '// エラー発生時はロールバック処理
    con_rollbacktrans
    Debug.Print "エラーが発生したのでロールバック処理を実施しました。"
    MsgBox "エラー番号:" & Err.Number & Chr(13) & "エラー内容:" & Err.Description
    Debug.Print "excute_sql:プログラムを終了します。"
    End

End Sub


'//----------------------------------------------------------------------------
'// 機能    ：実行したSQLで取得したレコードセットをExcelに貼り付け
'// 備考    ：is_filed_ ：Trueでフィールド名も書き込み
'//----------------------------------------------------------------------------
Private Sub paste_recordset(ByVal sheet_name_ As String, data_start_row_ As Long, data_start_col_ As Long, is_filed_ As Boolean)

    Dim i As Long

    If is_filed_ = True Then

        'フィールド名の書き出し
        For i = 0 To rs.Fields.count - 1
            ThisWorkbook.Worksheets(sheet_name_).Cells(data_start_row_, data_start_col_ + i).Value = rs.Fields(i).Name
        Next i
        
        data_start_row_ = data_start_row_ + 1
    
    End If
    
    'CopyFromRecordsetメソッドで基準セルを指定してデータの書き出し
    ThisWorkbook.Worksheets(sheet_name_).Cells(data_start_row_, data_start_col_).CopyFromRecordset rs

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：SQLを実行し、続けて取得したレコードセットをExcelに貼り付け
'// 備考    ：is_filed_ ：Trueでフィールド名も書き込み
'//----------------------------------------------------------------------------
Public Sub excute_sql_and_paste_recordset(ByVal str_sql_ As String, sheet_name_ As String, data_start_row_ As Long, data_start_col_ As Long, is_filed_ As Boolean)

    excute_sql str_sql_
    paste_recordset sheet_name_, data_start_row_, data_start_col_, is_filed_

End Sub


'//----------------------------------------------------------------------------
'// 機能    ：トランザクション開始処理
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub con_begintrans()

    con.BeginTrans
    Debug.Print "トランザクション開始"
    
End Sub

'//----------------------------------------------------------------------------
'// 機能    ：コミット処理（DBへの書き込み完了）
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub con_committrans()
    
    con.CommitTrans
    Debug.Print "コミット処理実施"

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：ロールバック処理
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub con_rollbacktrans()
    
    con.RollbackTrans
    Debug.Print "ロールバック処理実施"

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：DBへの接続終了処理
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub close_connection()

    On Error Resume Next

    con.Close
    rs.Close
    
    Set con = Nothing
    Set rs = Nothing
    
    On Error GoTo 0 '// エラー処理の命令取り消し
    Debug.Print "DBへの切断完了"

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：INSERT もしくは UPDATE のSQLを作成
'// 備考    ：引数は下記
'//           テーブル名、シート名、SQLタイプ指定行、データベースカラム名指定行
'//           データ格納開始行、データ格納開始列、データ格納終了列
'// 　　      （SQLタイプ：INWERT or UPDATE or WHERE）
'//----------------------------------------------------------------------------
Public Function make_sql(table_name As String, sheet_name As String, sql_type_defined_row As Long, db_col_name_defined_row As Long _
    , data_start_row As Long, data_start_col As Long, data_end_col As Long) As String

    Dim i As Long, j As Long
    Dim start_row As Long
    Dim start_col As Long, end_col As Long
    
    Dim sql_mode As Long    '//1:エラー終了　2:INSERT　3:UPDATE
    Dim sqltype As String
    Dim sql As String, sql_1 As String, sql_2 As String
    
    start_row = data_start_row

    start_col = data_start_col
    end_col = data_end_col
    
    sql_mode = 1    '// SQLモードが指定されなければエラー終了
    sqltype = ThisWorkbook.Worksheets(sheet_name).Cells(sql_type_defined_row, start_col).Value
    Select Case sqltype
        Case "INSERT"
            sql_mode = 2

        Case "UPDATE"
            sql_mode = 3
    
        Case "WHERE"
            sql_mode = 3
    End Select
    
    i = start_row
    If sql_mode = 2 Then    '//1:エラー終了　2:INSERT　3:UPDATE
    
        '// INSERT文の生成
        
        '// SQLを初期化
        sql = ""
        sql_1 = "INSERT INTO " & table_name & " ("
        sql_2 = "VALUES ("
    
        For j = start_col To end_col
    
            '// DBカラム名を追加
            sql_1 = sql_1 & ThisWorkbook.Worksheets(sheet_name).Cells(db_col_name_defined_row, j).Value
            If j <> end_col Then
                sql_1 = sql_1 & ", "
            Else
                sql_1 = sql_1 & ") "
            End If
            
            '// DBカラム名に対する値を追加
            sql_2 = sql_2 & "'" & ThisWorkbook.Worksheets(sheet_name).Cells(i, j).Value & "'"
            If j <> end_col Then
                sql_2 = sql_2 & ", "
            Else
                sql_2 = sql_2 & ") "
            End If
            
        Next
        
        sql = sql_1 + sql_2
                
    ElseIf sql_mode = 3 Then
    
        '// UPDATE文の生成
                    
        '// SQLを初期化
        sql = ""
        sql_1 = "UPDATE " & table_name & " SET "
        sql_2 = "WHERE "
    
        For j = start_col To end_col
    
            If ThisWorkbook.Worksheets(sheet_name).Cells(sql_type_defined_row, j).Value = "UPDATE" Then
    
                '// UPDATEするDBカラム名と値を追加
                sql_1 = sql_1 & ThisWorkbook.Worksheets(sheet_name).Cells(db_col_name_defined_row, j).Value & " = '" _
                     & ThisWorkbook.Worksheets(sheet_name).Cells(i, j).Value & "'"
                If j <> end_col Then
                    sql_1 = sql_1 & ", "
                Else
                    sql_1 = sql_1 & " "
                End If
                                
            ElseIf ThisWorkbook.Worksheets(sheet_name).Cells(sql_type_defined_row, j).Value = "WHERE" Then
            
                '// WHEREに対するDBカラム名と値を追加
                sql_2 = sql_2 & ThisWorkbook.Worksheets(sheet_name).Cells(db_col_name_defined_row, j).Value & " = '" _
                     & ThisWorkbook.Worksheets(sheet_name).Cells(i, j).Value & "'"
                If j <> end_col Then
                    sql_2 = sql_2 & " AND "
                Else
                    sql_2 = sql_2 & " "
                End If
            
            Else
                Debug.Print ("UPDATE文に対して、SQLタイプ設定が正しくありません。設定を見直して下さい。")
                make_sql = ""
                Exit Function
            End If
        Next
        
        sql = Replace(sql_1 + sql_2, ", WHERE", " WHERE")
        If Right(sql, 1) = "," Then sql = Mid(sql, 1, Len(sql) - 1)
        If Right(sql, 5) = " AND " Then sql = Mid(sql, 1, Len(sql) - 5)
    
    Else
        Debug.Print ("SQL生成モードに対して、SQLタイプ設定が正しくありません。" & Chr(13) & "プログラムを見直して下さい。")
        make_sql = ""
        Exit Function
    End If
    
    make_sql = sql
        
End Function

'//----------------------------------------------------------------------------
'// 機能    ：INSERT もしくは UPDATE のSQLを作成して実行
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub make_and_excute_sql(table_name As String, sheet_name As String, sql_type_defined_row As Long, db_col_name_defined_row As Long _
    , data_start_row As Long, data_start_col As Long, data_end_col As Long)

    Dim sql As String

   '// トランザクション開始
    con_begintrans

    '// SQL作成
    sql = make_sql(table_name, sheet_name, sql_type_defined_row, db_col_name_defined_row, data_start_row, data_start_col, data_end_col)
    
    '// SQL実行
    excute_sql sql
    
    '// コミット処理
    con_committrans
    
End Sub

'//----------------------------------------------------------------------------
'// 機能    ：連続してINSERT もしくは UPDATE のSQLを作成して実行
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub make_and_excute_sqls(table_name As String, sheet_name As String, sql_type_defined_row As Long, db_col_name_defined_row As Long _
    , data_start_row As Long, data_start_col As Long, data_end_col As Long)

    Dim i As Long
    Dim sql As String
    Dim val
    Dim data_end_row As Long

    '// データの最終行を検索
    i = 1
    Do While ThisWorkbook.Worksheets(sheet_name).Cells(data_start_row + i, data_start_col) <> ""
        i = i + 1
    Loop
    
    data_end_row = data_start_row + i - 1

    '// トランザクション開始
    con_begintrans
        
    For i = data_start_row To data_end_row

        '// SQL作成
        sql = make_sql(table_name, sheet_name, sql_type_defined_row, db_col_name_defined_row, i, data_start_col, data_end_col)
        '// SQL実行
        excute_sql sql
    
    Next
    
    '// コミット処理
    con_committrans
    
End Sub




