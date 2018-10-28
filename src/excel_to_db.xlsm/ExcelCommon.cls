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
'// ExcelCommonクラス
'//----------------------------------------------------------------------------


'// ↓関数情報コメント用のフォーマット
'//----------------------------------------------------------------------------
'// 関数名  ：
'// 機能    ：
'// 引数    ：
'// 戻り値  ：
'// 備考    ：
'//----------------------------------------------------------------------------

Option Explicit

Private book As Workbook
Private opened_book As Workbook

'//----------------------------------------------------------------------------
'// 関数名  ：
'// 機能    ：コンストラクタ
'// 引数    ：
'// 戻り値  ：
'// 備考    ：
'//----------------------------------------------------------------------------
Private Sub Class_Initialize()

End Sub

'//----------------------------------------------------------------------------
'// 関数名  ：
'// 機能    ：ワークブックのセット。インスタンス化後に使用する。
'// 引数    ：
'// 戻り値  ：
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub set_workbook(ByVal workbook_ As Workbook)

    Set book = workbook_

End Sub

'//----------------------------------------------------------------------------
'// 関数名  ：
'// 機能    ：指定したシートとセル範囲のデータをクリア
'// 引数    ：
'// 戻り値  ：
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub data_clear(ByVal sheet As String, start_row As Long, start_col As Long, end_row As Long, end_col As Long)
    
    book.Worksheets(sheet).Activate
    book.Worksheets(sheet).Range(Cells(start_row, start_col), Cells(end_row, end_col)).Clear

End Sub

'//----------------------------------------------------------------------------
'// 関数名  ：
'// 機能    ：指定したシートとセル範囲のデータをクリア
'// 引数    ：
'// 戻り値  ：
'// 備考    ：指定した行列から右側、下側に続く最大の範囲を計算してクリアする。
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
'// 関数名  ：
'// 機能    ：指定したセルから下側に続く最大のセル行を取得
'// 引数    ：
'// 戻り値  ：
'// 備考    ：
'//----------------------------------------------------------------------------
Public Function get_max_row_below(ByVal sheet As String, row As Long, col As Long) As Long

    If book Is Nothing Then
        MsgBox "Class ExcelCommon にワークブックがセットされていません。"
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
'// 関数名  ：
'// 機能    ：指定したセルから右側に続く最大のセル列を取得
'// 引数    ：
'// 戻り値  ：
'// 備考    ：
'//----------------------------------------------------------------------------
Public Function get_max_col_right(ByVal sheet As String, row As Long, col As Long) As Long

    If book Is Nothing Then
        MsgBox "Class ExcelCommon にワークブックがセットされていません。"
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
'// 関数名  ：
'// 機能    ：指定したファイルを開く
'// 引数    ：
'// 戻り値  ：
'// 備考    ：ファイル名はフォルダも含めて渡す
'//----------------------------------------------------------------------------
Public Sub open_file(ByVal file_name As String)

    Set opened_book = Workbooks.Open(file_name)
    Debug.Print file_name & "を開きました。"

End Sub

'//----------------------------------------------------------------------------
'// 関数名  ：
'// 機能    ：開いていたファイルを閉じる
'// 引数    ：
'// 戻り値  ：
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub close_opened_file(ByVal display_alerts As Boolean)

    Application.DisplayAlerts = display_alerts
    opened_book.Close
    Debug.Print "開いていたファイルを閉じました。"
    
    Application.DisplayAlerts = True

End Sub
