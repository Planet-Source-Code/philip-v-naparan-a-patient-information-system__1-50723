Attribute VB_Name = "Module1"
Option Explicit

Public file_path As String
Public patient() As String
Public curr_num As Integer
Public total_num As Integer

Public sDelFlag As String
Public sDelFlag_num As String

Public Function file_total() As Integer
Dim tmp As String
file_total = 0
Open file_path For Input As #2
    Do Until EOF(2)
        Line Input #2, tmp
        If tmp <> "" Then file_total = file_total + 1
    Loop
Close #2
tmp = ""
End Function
Public Sub load_file()
Dim c As Integer
Dim l As Integer
Dim array2 As Integer
Dim tmp1 As String
Dim tmp2 As String
Dim tmp3 As String
If file_total = 0 Then Exit Sub
ReDim patient(total_num, 6)
curr_num = 1
Open file_path For Input As #1
    For c = 1 To file_total
        Line Input #1, tmp1
        array2 = 1
        For l = 1 To Len(tmp1)
            tmp2 = Mid(tmp1, l, 1)
            If tmp2 = vbTab Then
                patient(c, array2) = tmp3
                array2 = array2 + 1
                tmp3 = ""
            ElseIf l = Len(tmp1) Then
                patient(c, array2) = tmp2
            Else
                tmp3 = tmp3 & tmp2
            End If
        Next l
    Next c
Close #1
c = 0
l = 0
tmp1 = ""
tmp2 = ""
tmp3 = ""
End Sub
Public Sub write_file(ByVal sTotal As Integer)
Dim c As Integer
Open file_path For Output As #3
    For c = 1 To sTotal
        Print #3, patient(c, 1) & vbTab & patient(c, 2) & vbTab & patient(c, 3) & vbTab & patient(c, 4) & vbTab & patient(c, 5) & vbTab & patient(c, 6)
    Next c
Close #3
c = 0
End Sub
Public Sub search_record(ByVal sSearchTXT As String, ByVal sTotal As Integer, ByRef sCur_Record As Integer)
Dim c As Integer
    For c = 1 To sTotal
        If LCase(patient(c, 1)) = LCase(sSearchTXT) Then
            sCur_Record = c
            MsgBox "Record found at record no. " & c, vbInformation, "Patient Info"
            Exit Sub
        End If
    Next c
MsgBox "The search for '" & sSearchTXT & "' has not been found in the record." & vbCrLf & "Please make sure it was exist in the record.", vbExclamation, "Patient Info"
End Sub
Public Function record_exist(ByVal sSearchTXT As String, ByVal sTotal As Integer) As Boolean
Dim c As Integer
    For c = 1 To sTotal
        If LCase(patient(c, 1)) = LCase(sSearchTXT) Then
            sDelFlag = patient(c, 6)
            record_exist = True
            sDelFlag_num = c
            Exit Function
        End If
    Next c
record_exist = False
End Function
