Attribute VB_Name = "Module2"
Option Explicit

Sub �f�[�^�ǉ�()
    Application.ScreenUpdating = False
    Dim i As Long
    Dim starttime As Double
    Dim endtime As Double
    Dim keika As Double
    
'    starttime = Timer
    
    With Sheets("�f�[�^�ǉ�")
        For i = 2 To .Cells(.Rows.Count, 1).End(xlUp).Row
            .Cells(i, 6) = Left(.Cells(i, 1), 4)
            .Cells(i, 7) = WorksheetFunction.VLookup( _
                .Cells(i, 2), Sheets("�}�X�^").Range("A:B"), 2, 0)
            
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
    
    'Call �N���A
End Sub


Sub �f�[�^�ǉ�4()
    ' ��ʍX�V���~
    Application.ScreenUpdating = False
    ' �v�Z���[�h���蓮�ɐݒ�
    Application.Calculation = xlCalculationManual
    
    ' �ϐ��̐錾
    Dim i As Long ' ���[�v�J�E���^
    Dim lastRow As Long ' �f�[�^�ǉ��V�[�g�̍ŏI�s
    Dim dataArray As Variant ' �f�[�^�ǉ��V�[�g�̃f�[�^���i�[����z��
    Dim masterArray As Variant ' �}�X�^�V�[�g�̃f�[�^���i�[����z��
    Dim dict As Object ' �}�X�^�f�[�^���i�[���鎫���I�u�W�F�N�g
    Set dict = CreateObject("Scripting.Dictionary") ' �����I�u�W�F�N�g�̏�����
    
'    ' �����J�n���Ԃ��L�^
'    starttime = Timer
    
    ' �}�X�^�V�[�g����f�[�^���擾
    With Sheets("�}�X�^")
        masterArray = .Range("A1:B" & .Cells(.Rows.Count, 1).End(xlUp).Row).Value
    End With
    
    ' �}�X�^�f�[�^�������Ɋi�[
    For i = 1 To UBound(masterArray, 1)
        dict(masterArray(i, 1)) = masterArray(i, 2)
    Next i
    
    ' �f�[�^�ǉ��V�[�g�̏���
    With Sheets("�f�[�^�ǉ�")
        ' �ŏI�s���擾
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        ' �f�[�^��z��Ɋi�[
        dataArray = .Range("A2:H" & lastRow).Value
        
        ' �e�s�̃f�[�^������
        For i = 1 To UBound(dataArray, 1)
            ' ��A�̑O4�������F�ɐݒ�
            dataArray(i, 6) = Left(dataArray(i, 1), 4)
            ' ��B�̒l���}�X�^�V�[�g���猟�����ė�G�ɐݒ�
            If dict.exists(dataArray(i, 2)) Then
                dataArray(i, 7) = dict(dataArray(i, 2))
            Else
                dataArray(i, 7) = "Not Found"
            End If
            
            ' ��E�̒l�Ɋ�Â��ė�H��"A"�܂���"B"��ݒ�
            If dataArray(i, 5) >= 50000 Then
                dataArray(i, 8) = "A"
            Else
                dataArray(i, 8) = "B"
            End If
        Next i
        
        ' ���������f�[�^���V�[�g�ɏ����߂�
        .Range("A2:H" & lastRow).Value = dataArray
    End With
    
    ' �v�Z���[�h�������ɖ߂�
    Application.Calculation = xlCalculationAutomatic
    ' ��ʍX�V���ĊJ
    Application.ScreenUpdating = True
    
End Sub

Sub �f�[�^�ǉ�3()
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
    
    With Sheets("�}�X�^")
        masterArray = .Range("A1:B" & .Cells(.Rows.Count, 1).End(xlUp).Row).Value
    End With
    
    For i = 1 To UBound(masterArray, 1)
        dict(masterArray(i, 1)) = masterArray(i, 2)
    Next i
    
    With Sheets("�f�[�^�ǉ�")
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


Sub �f�[�^�ǉ�2()
    Application.ScreenUpdating = False
    Dim indexx As Long
    Dim indexy As Long
        
    With Sheets("�f�[�^�ǉ�")
        indexx = .Cells(.Rows.Count, 1).End(xlUp).Row
        'indexy = .Cells(1, .Columns.Count).End(xlToLeft).Column
        indexy = 3
        
        
        
    End With

End Sub


Sub �N���A()
    'Range("F2", Range("F2").SpecialCells(xlLastCell)).Clear
    'Range("G2", Range("G2").SpecialCells(xlLastCell)).Clear
    'Range("H2", Range("H2").SpecialCells(xlLastCell)).Clear
    
    With Sheets("�f�[�^�ǉ�")
        .Range("F2:H" & .Cells(1, 1).End(xlDown).Row).Clear
    End With
    
End Sub



