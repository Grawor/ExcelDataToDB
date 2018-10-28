Option Explicit
'
' vbac.wsf 実行 Macro
'

Dim folder
dim file

'FileSystemObjectを生成
Dim FSO

folder = current_folder()
file = folder & "\temp.bat"
'MsgBox file

'バッチファイルを作成
Set FSO = CreateObject("Scripting.FileSystemObject")
With FSO.CreateTextFile(file)
    .WriteLine "cd " & folder
    .WriteLine "cscript vbac.wsf decombine"
    .Close
End With
Set FSO = Nothing


'バッチファイル実行
excute_bat(file)

'ファイルを削除
delete_bat(file)

MsgBox("vbac.wsf decombine done!")

'----- ここから関数 -----

Private Function current_folder()
    On Error Resume Next

    Dim objWshShell     ' WshShell オブジェクト

    Set objWshShell = WScript.CreateObject("WScript.Shell")
    current_folder = objWshShell.CurrentDirectory
    Set objWshShell = Nothing

End Function

Private Function excute_bat(ByVal file_name)

    ' WshShellオブジェクトを作成する
    Dim WshShell
    Set WshShell = WScript.CreateObject("WScript.Shell")
    
    ' batファイルを実行する
    WshShell.Run file_name,0,True
    
    ' オブジェクトを開放する
    Set WshShell = Nothing

End Function

Private Function delete_bat(ByVal file_name)

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(file_name) Then
        'WScript.Echo file_name & " file found."
    Else
        'WScript.Echo file_name & " file not found."
    End If

    fso.DeleteFile(file_name)

    If fso.FileExists(file_name) Then
        'WScript.Echo file_name & " file found."
    Else
        'WScript.Echo file_name & " file not found."
    End If

    Set fso = Nothing

End Function
