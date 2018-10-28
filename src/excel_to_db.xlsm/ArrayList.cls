VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//----------------------------------------------------------------------------
'// ArrayListクラス
'//----------------------------------------------------------------------------
Option Explicit
 
Private mvArray() As Variant
 
'//----------------------------------------------------------------------------
'// 関数名  ：Class_Initialize
'// 引数    ：なし
'// 戻り値  ：なし
'// 機能    ：コンストラクタ
'// 備考    ：
'//----------------------------------------------------------------------------
Private Sub Class_Initialize()
    ReDim mvArray(0)
End Sub
 
'//----------------------------------------------------------------------------
'// 関数名  ：Add
'// 機能    ：値をクラスに追加する
'// 引数    ：(I) asVal
'// 戻り値  ：なし
'// 備考    ：JavaのArrayList.Addと一緒
'//----------------------------------------------------------------------------
Public Sub add(asVal)
    On Error Resume Next
    
    Dim vArray()
    Dim iCnt
    
    iCnt = Me.count
    
    If (IsEmpty(mvArray(iCnt)) = True) Then
        mvArray(iCnt) = asVal
    Else
        ReDim Preserve mvArray(iCnt + 1)
        mvArray(iCnt + 1) = asVal
    End If
 
End Sub
 
'//----------------------------------------------------------------------------
'// 関数名  ：Remove
'// 機能    ：値をクラスから削除する
'// 引数    ：(I) asVal
'// 戻り値  ：なし
'// 備考    ：JavaのArrayList.Removeと一緒
'//----------------------------------------------------------------------------
Public Sub Remove(aiIndex)
    On Error Resume Next
    
    Dim vArray()
    Dim iCnt
    Dim i                       '// ループカウンタ
    Dim iAdjust                 '// mvArrayからvArrayへコピーするためのIndex
    
    iAdjust = 0
    iCnt = Me.count
    ReDim vArray(iCnt - 1)      '// 要素を削るのでその分もマイナスする
    
    For i = 0 To iCnt
        If (i <> aiIndex) Then
            vArray(iAdjust) = mvArray(i)
            iAdjust = iAdjust + 1
        End If
    Next
    
    Me.Clear
    ReDim mvArray(iAdjust)
    mvArray = vArray
 
End Sub
 
'//----------------------------------------------------------------------------
'// 関数名  ：GetVal
'// 機能    ：指定要素位置の値を取得する
'// 引数    ：(I) aiIndex
'// 戻り値  ：指定要素位置の値
'// 備考    ：JavaのArrayList.Getと一緒
'//----------------------------------------------------------------------------
Public Function GetVal(aiIndex)
    Dim vRet As Variant
    
    If (aiIndex > Me.count) Then
        vRet = Null
    Else
        vRet = mvArray(aiIndex)
    End If
    
    GetVal = vRet
End Function
 
'//----------------------------------------------------------------------------
'// 関数名  ：GetIndex
'// 機能    ：指定要素のリスト内で最初に検出された位置のインデックスを取得する
'// 引数    ：(I) asVal
'// 戻り値  ：指定要素のインデックス。要素がない場合は-1を返す。
'// 備考    ：JavaのArrayList.indexOfと一緒
'//----------------------------------------------------------------------------
Public Function GetIndex(asVal)
    On Error Resume Next
    
    Dim iCnt
    Dim i                       '// ループカウンタ
    Dim bExistFlg As Boolean    '// 要素存在フラグ（True：要素あり、False：なし）
    Dim iIndex
    
    bExistFlg = False
    iCnt = Me.count
    
    For i = 0 To iCnt
        If (mvArray(i) = asVal) Then
            bExistFlg = True
            Exit For
        End If
    Next
    
    If (bExistFlg = True) Then
        iIndex = i
    Else
        iIndex = -1
    End If
    
    GetIndex = iIndex
End Function
 
'//----------------------------------------------------------------------------
'// 関数名  ：Clear
'// 機能    ：配列要素をクリアする
'// 引数    ：なし
'// 戻り値  ：なし
'// 備考    ：JavaのArrayList.Clearと一緒
'//----------------------------------------------------------------------------
Public Function Clear()
    Call Class_Initialize
End Function
 
'//----------------------------------------------------------------------------
'// 関数名  ：Count
'// 機能    ：配列要素数を調べる
'// 引数    ：なし
'// 戻り値  ：Long ：マッピング数
'// 備考    ：JavaのArrayList.Countと一緒
'//----------------------------------------------------------------------------
Public Function count() As Long
    count = UBound(mvArray)
End Function
 
'//----------------------------------------------------------------------------
'// 関数名  ：Contains
'// 機能    ：引数値がリストに含まれるかを確認する
'// 引数    ：(I) asVal
'// 戻り値  ：Boolean ：True：引数がリストに含まれる、False：含まれない
'// 備考    ：JavaのArrayList.containsと一緒
'//----------------------------------------------------------------------------
Public Function Contains(asVal) As Boolean
    Dim iCnt As Long
    Dim iListCnt As Long
    Dim bRet As Boolean
    
    bRet = False
    iListCnt = Me.count
    
    Do While (iCnt <= iListCnt)
        If (asVal = mvArray(iCnt)) Then
            bRet = True
        Exit Do
    End If
    
    iCnt = iCnt + 1
    Loop
    
    Contains = bRet
End Function
 
'//----------------------------------------------------------------------------
'// 関数名  ：ContainsU
'// 機能    ：引数値がリストに含まれるかを確認する。大文字で統一して比較。
'// 引数    ：(I) asVal
'// 戻り値  ：Boolean ：True：引数がリストに含まれる、False：含まれない
'// 備考    ：JavaのArrayList.containsと一緒
'//----------------------------------------------------------------------------
Public Function ContainsU(asVal) As Boolean
    Dim iCnt As Long
    Dim iListCnt As Long
    Dim bRet As Boolean
    
    bRet = False
    iListCnt = Me.count
    
    Do While (iCnt <= iListCnt)
    If (UCase(asVal) = UCase(mvArray(iCnt))) Then
        bRet = True
        Exit Do
    End If
    
    iCnt = iCnt + 1
    Loop
    
    ContainsU = bRet
End Function
 
'//----------------------------------------------------------------------------
'// 関数名  ：ContainsInstr
'// 機能    ：引数文字列値がリスト文字列の一部に含まれるかを確認する。大文字で統一して比較。
'// 引数    ：(I) asVal
'// 戻り値  ：Boolean ：True：引数がリストに含まれる、False：含まれない
'// 備考    ：JavaのArrayList.containsと一緒
'//----------------------------------------------------------------------------
Public Function ContainsInStr(asVal) As Boolean
    Dim iCnt As Long
    Dim iListCnt As Long
    Dim bRet As Boolean
    
    bRet = False
    iListCnt = Me.count
    
    Do While (iCnt <= iListCnt)
        If (InStr(1, UCase(asVal), UCase(mvArray(iCnt))) > 0) Then
            bRet = True
            Exit Do
        End If
    
        iCnt = iCnt + 1
    Loop
    
    ContainsInStr = bRet
End Function
