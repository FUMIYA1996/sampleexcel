Attribute VB_Name = "Module1"
Option Explicit

Sub main()
    Dim starttime As Double
    starttime = Timer
    
    Sheets("データ追加").Range("K7").Value = Timer
    
    Call データ追加4
    
    Sheets("データ追加").Range("L7").Value = Timer
    
    'MsgBox Timer - starttime
    
    Call クリア
    
End Sub

Sub del()
    Call クリア
End Sub

