VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "����1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ShowResult_Click()
 '- �������� �� ������� ������ ����������� ------------------------------------
 For i = 1 To Sheets.Count
    If (Sheets(i).Name = "����������") Then
        MsgBox "������ ���� � ���������� �����������, �� ����� ������.", , _
                    "�����������"
        Application.DisplayAlerts = False
        Sheets(i).Delete
        Application.DisplayAlerts = True
    End If
 Next i
 '- ����������� ��������� -----------------------------------------------------
 Sheets.Add(After:=Sheets(Sheets.Count)).Name = "����������"
 Sheets("����������").Cells(1, 1).Value = "���"
 Sheets("����������").Cells(1, 2).Value = "�������"
 Sheets("����������").Cells(1, 3).Value = "���"
 Sheets("����������").Cells(1, 4).Value = "�����������"
 Sheets("����������").Cells(1, 5).Value = "���"
 Sheets("����������").Cells(1, 6).Value = "����������� �����������"
 Sheets("����������").Cells(1, 7).Value = "���� �����������"
 Sheets("����������").Cells(1, 8).Value = "���-�� �����������"
 '- ���������� ����������� ����� � ������ ����� -------------------------------
 Dim RowAtListA As Single
 Dim RowAtListB As Single
 Dim RowAtListV As Single
 Dim RowAtListG As Single
 Dim RowAtListD As Single
 RowAtListA = Sheets("������ �").UsedRange.Rows.Count
 RowAtListB = Sheets("������ �").UsedRange.Rows.Count
 RowAtListV = Sheets("������ �").UsedRange.Rows.Count
 RowAtListG = Sheets("������ �").UsedRange.Rows.Count
 RowAtListD = Sheets("������ �").UsedRange.Rows.Count
 '- ���������� �������� ������� � ������ --------------------------------------
 Dim TempListB() As String
 Call FillOutSheet(TempListB(), "������ �", RowAtListB)
 
 Dim TempListV() As String
 Call FillOutSheet(TempListV(), "������ �", RowAtListV)
 
 Dim TempListG() As String
 Call FillOutSheet(TempListG(), "������ �", RowAtListG)
 
 Dim TempListD() As String
 Call FillOutSheet(TempListD(), "������ �", RowAtListD)
 '- ��������� ������ � ����� "������ �" ---------------------------------------
 Dim Result() As String
 ReDim Result(RowAtListA, 11)
 For i = 1 To RowAtListA - 1
    '- ������ ������� ����� �������� � ������ ---------------------------------
    Dim temp() As String
    temp = Split(Sheets("������ �").Cells(i + 1, 1).Value, " ", 3)
    For j = 1 To UBound(temp) - LBound(temp) + 1
        Result(i - 1, j - 1) = temp(j - 1)
    Next j
    Erase temp
    '- ������ ��� � ��� � ������ ----------------------------------------------
    temp = Split(Sheets("������ �").Cells(i + 1, 2).Value, "-", 2)
    For j = 1 To UBound(temp) - LBound(temp) + 1
        Result(i - 1, 2 + 2 * j) = temp(j - 1)
    Next j
    Erase temp
    '- ������ ���� � ������ ---------------------------------------------------
    temp = Split(Sheets("������ �").Cells(i + 1, 3).Value, ".", 3)
    For j = 1 To UBound(temp) - LBound(temp) + 1
        Result(i - 1, 7 + j) = temp(j - 1)
    Next j
    Erase temp
    '- ����� ������������ �� ��� � ������� ��� � ������ �������� � ������ -----
    Dim IOF As String
    IOF = Result(i - 1, 1) & " " & Result(i - 1, 2) & " " & Result(i - 1, 0)
    Call CompareAndPush(Result(), TempListB(), IOF, 3, i)
    '- ����� ������������ �� ��� � ������ ����������� � ������ ----------------
    Call CompareAndPush(Result(), TempListV(), Result(i - 1, 4), 5, i)
    '- ����� ������������ �� ��� � ������ ����������� ����������� � ������ ----
    Call CompareAndPush(Result(), TempListG(), Result(i - 1, 6), 7, i)
    '- ����� ������������ �� ���� � ������� ����-��-�� � ������ ���-�� --------
    '- ����������� � ������ ---------------------------------------------------
    Dim ggggmmdd As String
    ggggmmdd = Result(i - 1, 10) & "-" & Result(i - 1, 9) & "-" & _
                    Result(i - 1, 8)
    Call CompareAndPush(Result(), TempListD(), ggggmmdd, 11, i)
    '- ����������� �� ����� ���������� �� ������� -----------------------------
    Sheets("����������").Cells(i + 1, 1).Value = Result(i - 1, 0) & " " & _
                    Result(i - 1, 1) & " " & Result(i - 1, 2)
    For j = 3 To 7
    Sheets("����������").Cells(i + 1, j - 1).Value = Result(i - 1, j)
    Next j
    Sheets("����������").Cells(i + 1, 7).Value = Result(i - 1, 8) & "." & _
            Result(i - 1, 9) & "." & Result(i - 1, 10)
    Sheets("����������").Cells(i + 1, 8).Value = Result(i - 1, 11)
 Next i
 Erase TempListB
 Erase TempListV
 Erase TempListG
 Erase TempListD
 Erase Result
 For i = 1 To 8
    Sheets("����������").Cells(i + 1, i).EntireColumn.AutoFit
 Next i
End Sub
'- ������ ������� ��������� ������ ������� ������ ���� ������� � ��������� ----
'- ����� ----------------------------------------------------------------------
Sub FillOutSheet(TempSheet() As String, NameOfSheet As String, _
                    RowAtList As Single)
    ReDim TempSheet(RowAtList, 2)
    For i = 1 To RowAtList - 1
        TempSheet(i - 1, 0) = Sheets(NameOfSheet).Cells(i + 1, 1).Value
        TempSheet(i - 1, 1) = Sheets(NameOfSheet).Cells(i + 1, 2).Value
    Next i

End Sub
'- ������ ������� ���������� �������� �� ������ ������� � �������� ��������, --
'- ��� ���������� ���������� � ������ ������ � �������� ������� ---------------
Sub CompareAndPush(Result() As String, TempArr() As String, _
                    ByVal Condition As String, ByVal Position As Integer, _
                    ByVal i As Integer)
    For j = 1 To UBound(TempArr) - LBound(TempArr)
        If (TempArr(j - 1, 0) = Condition) Then
            Result(i - 1, Position) = TempArr(j - 1, 1)
            Exit For
        End If
    Next j
End Sub

