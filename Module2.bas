Attribute VB_Name = "Module2"
Option Explicit

Sub データ追加()
    Application.ScreenUpdating = False
    Dim i As Long
    Dim starttime As Double
    Dim endtime As Double
    Dim keika As Double
    
'    starttime = Timer
    
    With Sheets("データ追加")
        For i = 2 To .Cells(.Rows.Count, 1).End(xlUp).Row
            .Cells(i, 6) = Left(.Cells(i, 1), 4)
            .Cells(i, 7) = WorksheetFunction.VLookup( _
                .Cells(i, 2), Sheets("マスタ").Range("A:B"), 2, 0)
            
            If .Cells(i, 5) >= 50000 Then
                .Cells(i, 8) = "A"
            Else
                .Cells(i, 8) = "B"
            End If
        Next
    End With
            
'    endtime = Timer
    
'    keika = endtime - starttime
'    MsgBox keika
    
    'Call クリア
End Sub


Sub データ追加4()
    ' 画面更新を停止
    Application.ScreenUpdating = False
    ' 計算モードを手動に設定
    Application.Calculation = xlCalculationManual
    
    ' 変数の宣言
    Dim i As Long ' ループカウンタ
    Dim lastRow As Long ' データ追加シートの最終行
    Dim dataArray As Variant ' データ追加シートのデータを格納する配列
    Dim masterArray As Variant ' マスタシートのデータを格納する配列
    Dim dict As Object ' マスタデータを格納する辞書オブジェクト
    Set dict = CreateObject("Scripting.Dictionary") ' 辞書オブジェクトの初期化
    
'    ' 処理開始時間を記録
'    starttime = Timer
    
    ' マスタシートからデータを取得
    With Sheets("マスタ")
        masterArray = .Range("A1:B" & .Cells(.Rows.Count, 1).End(xlUp).Row).Value
    End With
    
    ' マスタデータを辞書に格納
    For i = 1 To UBound(masterArray, 1)
        dict(masterArray(i, 1)) = masterArray(i, 2)
    Next i
    
    ' データ追加シートの処理
    With Sheets("データ追加")
        ' 最終行を取得
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        ' データを配列に格納
        dataArray = .Range("A2:H" & lastRow).Value
        
        ' 各行のデータを処理
        For i = 1 To UBound(dataArray, 1)
            ' 列Aの前4文字を列Fに設定
            dataArray(i, 6) = Left(dataArray(i, 1), 4)
            ' 列Bの値をマスタシートから検索して列Gに設定
            If dict.exists(dataArray(i, 2)) Then
                dataArray(i, 7) = dict(dataArray(i, 2))
            Else
                dataArray(i, 7) = "Not Found"
            End If
            
            ' 列Eの値に基づいて列Hに"A"または"B"を設定
            If dataArray(i, 5) >= 50000 Then
                dataArray(i, 8) = "A"
            Else
                dataArray(i, 8) = "B"
            End If
        Next i
        
        ' 処理したデータをシートに書き戻す
        .Range("A2:H" & lastRow).Value = dataArray
    End With
    
    ' 計算モードを自動に戻す
    Application.Calculation = xlCalculationAutomatic
    ' 画面更新を再開
    Application.ScreenUpdating = True
    
End Sub

Sub データ追加3()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Dim i As Long
    Dim starttime As Double
    Dim endtime As Double
    Dim keika As Double
    Dim lastRow As Long
    Dim dataArray As Variant
    Dim masterArray As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
'    starttime = Timer
    
    With Sheets("マスタ")
        masterArray = .Range("A1:B" & .Cells(.Rows.Count, 1).End(xlUp).Row).Value
    End With
    
    For i = 1 To UBound(masterArray, 1)
        dict(masterArray(i, 1)) = masterArray(i, 2)
    Next i
    
    With Sheets("データ追加")
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        dataArray = .Range("A2:H" & lastRow).Value
        
        For i = 1 To UBound(dataArray, 1)
            dataArray(i, 6) = Left(dataArray(i, 1), 4)
            If dict.exists(dataArray(i, 2)) Then
                dataArray(i, 7) = dict(dataArray(i, 2))
            Else
                dataArray(i, 7) = "Not Found"
            End If
            
            If dataArray(i, 5) >= 50000 Then
                dataArray(i, 8) = "A"
            Else
                dataArray(i, 8) = "B"
            End If
        Next i
        
        .Range("A2:H" & lastRow).Value = dataArray
    End With
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
'    endtime = Timer
'    keika = endtime - starttime
'    MsgBox keika
End Sub


Sub データ追加2()
    Application.ScreenUpdating = False
    Dim indexx As Long
    Dim indexy As Long
        
    With Sheets("データ追加")
        indexx = .Cells(.Rows.Count, 1).End(xlUp).Row
        'indexy = .Cells(1, .Columns.Count).End(xlToLeft).Column
        indexy = 3
        
        
        
    End With

End Sub


Sub クリア()
    'Range("F2", Range("F2").SpecialCells(xlLastCell)).Clear
    'Range("G2", Range("G2").SpecialCells(xlLastCell)).Clear
    'Range("H2", Range("H2").SpecialCells(xlLastCell)).Clear
    
    With Sheets("データ追加")
        .Range("F2:H" & .Cells(1, 1).End(xlDown).Row).Clear
    End With
    
End Sub



