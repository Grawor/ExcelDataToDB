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
'// 機能    ：
'// 備考    ：
'//----------------------------------------------------------------------------

Option Explicit

Private book As Workbook
Private opened_book As Workbook

'//----------------------------------------------------------------------------
'// 機能    ：コンストラクタ
'// 備考    ：
'//----------------------------------------------------------------------------
Private Sub Class_Initialize()

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：ワークブックのセット。インスタンス化後に使用する。
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub set_workbook(ByVal workbook_ As Workbook)

    Set book = workbook_

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：指定したシートとセル範囲のデータをクリア
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub data_clear(ByVal sheet As String, start_row As Long, start_col As Long, end_row As Long, end_col As Long)

    If book Is Nothing Then
        MsgBox "Class ExcelCommon にワークブックがセットされていません。"
        Exit Sub
    End If

    book.Worksheets(sheet).Activate
    book.Worksheets(sheet).Range(Cells(start_row, start_col), Cells(end_row, end_col)).Clear

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：指定したシートとセル範囲のデータをクリア
'// 備考    ：指定した行列から右側、下側に続く最大の範囲を計算してクリアする。
'//----------------------------------------------------------------------------
Public Sub data_clear2(ByVal sheet As String, start_row As Long, start_col As Long)

    If book Is Nothing Then
        MsgBox "Class ExcelCommon にワークブックがセットされていません。"
        Exit Sub
    End If

    Dim end_row As Long
    Dim end_col As Long
    
    end_row = Me.get_max_row_below(sheet, start_row, start_col)
    end_col = Me.get_max_col_right(sheet, start_row, start_col)
    
    book.Worksheets(sheet).Activate
    book.Worksheets(sheet).Range(Cells(start_row, start_col), Cells(end_row, end_col)).Clear

End Sub


'//----------------------------------------------------------------------------
'// 機能    ：指定したセルから下側に続く最大のセル行を取得
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
'// 機能    ：指定したセルから右側に続く最大のセル列を取得
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
'// 機能    ：指定したセルから下側に続く最大のセルに対して書式を変更
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub change_format_max_row_below(ByVal sheet As String, row As Long, col As Long, format_ As String)

    If book Is Nothing Then
        MsgBox "Class ExcelCommon にワークブックがセットされていません。"
        Exit Sub
    End If

    Dim val As String
    Dim i As Long
    i = 0
    Do While book.Worksheets(sheet).Cells(row + i, col) <> ""
        book.Worksheets(sheet).Cells(row + i, col).NumberFormatLocal = "@"
        val = format(book.Worksheets(sheet).Cells(row + i, col).Value, format_)
        book.Worksheets(sheet).Cells(row + i, col) = val
        i = i + 1
    Loop
    
End Sub

'//----------------------------------------------------------------------------
'// 機能    ：指定したセル範囲のレコードを削除
'// 備考    ：end_col = 0 の場合は start_col から右側に続く最大のセル列を自動指定
'//----------------------------------------------------------------------------
Public Sub delete_cells(ByVal sheet As String, row As Long, start_col As Long, end_col As Long)

    If book Is Nothing Then
        MsgBox "Class ExcelCommon にワークブックがセットされていません。"
        Exit Sub
    End If

    If end_col = 0 Then
    
        end_col = get_max_col_right(sheet, row, start_col)
    
    End If
    
    With book.Worksheets(sheet)
    
        .Range(.Cells(row, start_col), .Cells(row, end_col)).Delete (xlShiftUp)

    End With

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：セルの値が指定したカラムの中に含まれるか判定する。
'// 備考    ：
'//----------------------------------------------------------------------------
Public Function exist_val(ByVal val, sheet As String, row As Long, col As Long) As Boolean

    If book Is Nothing Then
        MsgBox "Class ExcelCommon にワークブックがセットされていません。"
        Exit Function
    End If

    Dim i As Long
    i = 0
    Do While book.Worksheets(sheet).Cells(row + i, col) <> ""
        
        If val = book.Worksheets(sheet).Cells(row + i, col).Value Then
            exist_val = True
            Exit Function
        End If
        
        i = i + 1
    Loop

    exist_val = False

End Function


'//----------------------------------------------------------------------------
'// 機能    ：指定したファイルを開く。
'// 備考    ：ファイル名はフォルダも含めて渡す。
'//----------------------------------------------------------------------------
Public Sub open_file(ByVal file_name As String)

    Set opened_book = Workbooks.Open(file_name)
    Debug.Print file_name & "を開きました。"

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：開いていたファイルを閉じる。
'// 備考    ：
'//----------------------------------------------------------------------------
Public Sub close_opened_file(ByVal display_alerts As Boolean)

    Application.DisplayAlerts = display_alerts
    opened_book.Close
    Debug.Print "開いていたファイルを閉じました。"
    
    Application.DisplayAlerts = True

End Sub

'//----------------------------------------------------------------------------
'// 機能    ：画面描画などを停止して実行を早くする。
'// 備考    ：Focus = True '描画停止、イベント抑制、手動計算
'//           Focus = False '描画再開、イベント監視再開、自動計算
'//----------------------------------------------------------------------------
Sub focus(ByVal Flag As Boolean)
    With Application
        .EnableEvents = Not Flag
        .ScreenUpdating = Not Flag
        .Calculation = IIf(Flag, xlCalculationManual, xlCalculationAutomatic)
    End With
End Sub
