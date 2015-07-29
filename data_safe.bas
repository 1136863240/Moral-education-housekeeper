Attribute VB_Name = "data_safe"
Function d_en(data As String) As String
    Dim data_result As String, data_temp As String
    Dim data_asc() As Integer, data_index As Integer
    ReDim data_asc(Len(data) - 1)
    For data_index = 0 To Len(data) - 1
        data_asc(data_index) = Asc(Mid(data, data_index + 1, 1))
        Dim flag As Boolean
        If data_asc(data_index) Mod 2 = 1 Then flag = True Else flag = False
        data_asc(data_index) = data_asc(data_index) * 10
        If flag = True Then data_asc(data_index) = data_asc(data_index) + 1
    Next
    For data_index = 0 To Len(data) - 1
        data_result = data_result & Chr(data_asc(data_index))
    Next
    d_en = data_result
End Function

Function d_de(data As String) As String
    Dim data_result As String, data_temp As String
    Dim data_asc() As Integer, data_index As Integer
    ReDim data_asc(Len(data) - 1)
    For data_index = 0 To Len(data) - 1
        data_asc(data_index) = Asc(Mid(data, data_index + 1, 1))
        If (data_asc(data_index) - 1) Mod 10 = 0 Then
            data_asc(data_index) = data_asc(data_index) - 1 / 10
        Else
            data_asc(data_index) = data_asc(data_index) / 10
        End If
    Next
    For data_index = 0 To Len(data) - 1
        data_result = data_result & Chr(data_asc(data_index))
    Next
    d_de = data_result
End Function
