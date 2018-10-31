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
'// ExcelDataGetterクラス
'//----------------------------------------------------------------------------

Option Explicit

'//----------------------------------------------------------------------------
'// 【使用例】
'//  ①インスタンス生成
'//　②set_info：データを読み込むExcel情報をインプット
'//  [データ取得に時間が掛かるならば、この間に読込ファイルを開く処理を入れる（別途）]
'//　③define_data_rows：データを読み込む行の開始から終了行を決定
'//　④set_data_array：特定の列に対するデータを一時的に格納
'//　⑤paste_data_to_excel：指定したExcelファイルのセルに格納データを貼り付け
'//----------------------------------------------------------------------------


Private folder_name As String
Private file_name As String
Private sheet_name As String
Private target As String            '// フォルダ名 + "[" + ファイル名 + "]" + シート名 + "!"

Private data_first_row As Long      '// 別Excelのデータが格納されている最初の行
Private data_last_row As Long       '// 別Excelのデータが格納されている最後の行
Private data() As Variant           '// 別Excelのデータを格納する配列

'//----------------------------------------------------------------------------
'// 機能    ：コンストラクタ
'// 備考    ：
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
'// 機能    ：データを読み込む対象Excelのフォルダ、ファイル名、シート名を設定
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub set_info(ByVal folder_name_ As String, file_name_ As String, sheet_name_ As String)

    Dim buf As String
    Dim open_file_name As String

    folder_name = folder_name_
    file_name = file_name_
    sheet_name = sheet_name_

    ''対象ブックを選択します
    '// open_file_name = Application.GetOpenFilename("Microsoft Excel ブック,*.xls")
    open_file_name = folder_name_ & "\" & file_name_
    
    '// ファイルが存在しなければ終了
    If Dir(open_file_name) = "" Then
        MsgBox (open_file_name & " が存在しません。" & Chr(13) _
            & "フォルダ名とファイル名を確認して下さい。")
        Exit Sub
    End If
    
    ''ファイル名に[]を付ける
    open_file_name = Replace(open_file_name, Dir(open_file_name), "[" & Dir(open_file_name) & "]")
    'Debug.Print (open_file_name)
    
    ''対象ワークシート名を取得
    '// sheet_name = InputBox("読み込むワークシート名を入力してください。")
    target = "'" & open_file_name & sheet_name & "'!"
    
    ''ワークシート名が正しいかどうか、まず読み込んでみる
    On Error Resume Next
    buf = ExecuteExcel4Macro(target & "R1C1")
    If Err <> 0 Then
        MsgBox "ワークシート [ " & sheet_name & " ] を読めませんでした。", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0 '// エラー処理の命令取り消し

    Debug.Print target

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：データが入っている開始行・終了行を特定
'// 備考    ：カラム名が存在するならば、データ検索開始行から下側に検索を進めて最終行を特定する。
'// 　　　　　※下側への検索時、空白数が連続して設定数未満であれば検索を進める。
'//----------------------------------------------------------------------------
Public Sub define_data_rows(ByVal field_name_ As String, field_row_ As Long, field_col_ As Long, data_first_row_ As Long, serch_end_blank_count_ As Long)

    Dim i As Long
    Dim buf As String
    Dim blank_count As Long

    '// カラム名が存在しなければ終了
    If is_existed_field_name(field_name_, field_row_, field_col_) = False Then
        MsgBox "Excel台帳にカラム名 " & field_name_ & " が存在しません。"
        Exit Sub
    End If

    ''データの読み込み
    i = 0
    blank_count = 0
    'buf = get_cell_value(data_first_row_ + i, field_col_)
    'If buf = "0" Then blank_count = blank_count + 1 Else blank_count = 0    '// セルが空白ならブランクカウントに+1
    
    Do While blank_count < serch_end_blank_count_
        
        '// データを読込
        buf = ExecuteExcel4Macro(target & "R" & data_first_row_ + i & "C" & field_col_)
        
        If buf = "0" Then blank_count = blank_count + 1 Else blank_count = 0    '// セルが空白ならブランクカウントに+1
        
        i = i + 1
    Loop

    data_first_row = data_first_row_
    data_last_row = data_first_row_ + i - blank_count - 1
    If data_last_row < data_first_row Then data_last_row = data_first_row
    
    Debug.Print "検索による　データ開始：" & data_first_row & "　データ終了：" & data_last_row

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：指定したExceデータを配列に格納して返す関数
'// 備考    ：①set_data_array ②paste_data_to_excel をセットで使う場合が多い。
'//----------------------------------------------------------------------------
Public Function set_data_array(ByVal field_name_ As String, field_row_ As Long, field_col_ As Long) As Variant()

    Dim i As Long
    Dim buf As String
     
    '// カラム名が存在しなければ終了
    If is_existed_field_name(field_name_, field_row_, field_col_) = False Then
        MsgBox "Excel台帳にカラム名 " & field_name_ & " が存在しません。"
        ReDim data(0)
        Exit Function
    End If

    '// 格納データをクリア
    Call Me.clear_data
    
    ''データの読み込み
    For i = data_first_row To data_last_row
    
        buf = get_cell_value(i, field_col_)
        Me.add_data buf
    
    Next
    
    set_data_array = data
    
End Function

'//----------------------------------------------------------------------------
'// 機能    ：指定したExcelファイルの特定セルに格納していたデータをセット
'// 備考    ：①set_data_array ②paste_data_to_excel をセットで使う場合が多い。
'//----------------------------------------------------------------------------
Public Sub paste_data_to_excel(ByVal workbook_ As Workbook, sheet_name_ As String, row_ As Long, col_ As Long)

    Dim i As Long

    For i = 0 To Me.count_data
    
        workbook_.Worksheets(sheet_name_).Cells(row_ + i, col_) = Me.get_data_val(i)
        Debug.Print "i：" & i & "　　行：" & row_ + i & "　　値：" & Me.get_data_val(i)
        
    Next
    Debug.Print ""

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：別Excelのカラム名が存在するか判定する関数
'// 備考    ：Excelの行列が変わっていないかどうかのチェック用
'//----------------------------------------------------------------------------
Private Function is_existed_field_name(ByVal field_name_ As String, row_ As Long, col_ As Long) As Boolean

    Dim buf As String
    
    buf = get_cell_value(row_, col_)
    
    If InStr(buf, field_name_) > 0 Then is_existed_field_name = True _
        Else is_existed_field_name = False

End Function

'//----------------------------------------------------------------------------
'// 機能    ：別Excelの指定した行列に対して、セル値を返す関数
'// 備考    ：
'//----------------------------------------------------------------------------
Private Function get_cell_value(ByVal row_ As Long, ByVal col_ As Long) As String

    get_cell_value = ExecuteExcel4Macro(target & "R" & row_ & "C" & col_)

End Function

'//----------------------------------------------------------------------------
'// 機能    ：データの値を配列に追加
'// 備考    ：
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
'// 機能    ：指定要素位置の値を取得する
'// 備考    ：JavaのArrayList.Getと一緒
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
'// 機能    ：データ配列の要素数を調べる。
'// 備考    ：
'//----------------------------------------------------------------------------
Public Function count_data() As Long
    count_data = UBound(data)
End Function

'//----------------------------------------------------------------------------
'// 機能    ：データ配列をクリアする。
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub clear_data()
    ReDim data(0)
End Sub
