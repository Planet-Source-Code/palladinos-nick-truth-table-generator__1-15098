Attribute VB_Name = "Module2"
Public Function character_next(expr As String, pos As Byte) As Boolean
       If character(Mid(expr, pos + 1, 1)) Or _
          InStr(1, "+*)", Mid(expr, pos + 1, 1)) <> 0 Or _
          pos = Len(expr) Then
          character_next = True
       Else
          character_next = False
       End If
End Function
Public Function character_prev(expr As String, pos As Byte) As Boolean
       If pos = 1 Then
          character_prev = True: Exit Function
       End If
       If character(Mid(expr, pos - 1, 1)) Or _
          InStr(1, "+*~(", Mid(expr, pos - 1, 1)) <> 0 Then
          character_prev = True
       Else
          character_prev = False
       End If
End Function
Public Function zeroone_next(expr As String, pos As Byte) As Boolean
       If InStr(1, "+*)", Mid(expr, pos + 1, 1)) <> 0 Or _
          pos = Len(expr) Then
          zeroone_next = True
       Else
          zeroone_next = False
       End If
End Function
Public Function zeroone_prev(expr As String, pos As Byte) As Boolean
       If pos = 1 Then
          zeroone_prev = True: Exit Function
       End If
       If InStr(1, "+*~(", Mid(expr, pos - 1, 1)) <> 0 Then
          zeroone_prev = True
       Else
          zeroone_prev = False
       End If
End Function
Public Function not_prev(expr As String, pos As Byte) As Boolean
       If pos = 1 Then
          not_prev = True
       Else
          If InStr(1, "+*(", Mid(expr, pos - 1, 1)) <> 0 Then
             not_prev = True
          Else
             not_prev = False
          End If
       End If
End Function

Public Function not_next(expr As String, pos As Byte) As Boolean
       If character(Mid(expr, pos + 1, 1)) _
          Or zeroone_validation(Mid(expr, pos + 1, 1)) _
          Or Mid(expr, pos + 1, 1) = "(" Then
          not_next = True
       Else
          not_next = False
       End If
End Function



Public Function andor_prev(expr As String, pos As Byte) As Boolean
       If pos = 1 Then
          andor_prev = False
       Else
          If character(Mid(expr, pos - 1, 1)) Or zeroone_validation(Mid(expr, pos - 1, 1)) _
             Or Mid(expr, pos - 1, 1) = ")" Then
             andor_prev = True
          Else
             andor_prev = False
          End If
       End If
End Function

Public Function andor_next(expr As String, pos As Byte) As Boolean
       If Mid(expr, pos + 1, 1) = "~" Then
          If character(Mid(expr, pos + 2, 1)) Or zeroone_validation(Mid(expr, pos + 2, 1)) _
             Or Mid(expr, pos + 2, 1) = "(" Then
             andor_next = True
          Else
             andor_next = False
          End If
       Else
          If character(Mid(expr, pos + 1, 1)) Or zeroone_validation(Mid(expr, pos + 1, 1)) _
             Or Mid(expr, pos + 1, 1) = "(" Then
             andor_next = True
          Else
             andor_next = False
          End If
       End If
End Function




