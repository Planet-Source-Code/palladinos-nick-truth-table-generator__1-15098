Attribute VB_Name = "Module4"
Public Function valid_parenthesis(expr As String) As Boolean
       Dim i As Byte
       Dim opened As Byte
       Dim closed As Byte
       Dim flag As Boolean
       flag = True
       i = 0: opened = 0: closed = 0
       Do
            i = i + 1
            If i > Len(expr) Then
               Exit Do
            End If
            If Mid(expr, i, 1) = "(" Then
               opened = opened + 1
            Else
               If Mid(expr, i, 1) = ")" Then
                  closed = closed + 1
               End If
            End If
            If closed > opened Then
               valid_parenthesis = False: Exit Function
            End If
       Loop While flag
       If closed = opened Then
          valid_parenthesis = True
       Else
          valid_parenthesis = False
       End If
End Function
Public Function character(char As String) As Boolean
       Dim upchar As String
       upchar = UCase(char)
       If upchar >= "A" And upchar <= "Z" Then
          character = True
       Else
          character = False
       End If
End Function
Public Function zeroone_validation(expr As String) As Boolean
       If expr = "1" Or expr = "0" Then
          zeroone_validation = True
       Else
          zeroone_validation = False
       End If
End Function

Public Function open_next(expr As String, pos As Byte) As Boolean
       If Mid(expr, pos + 1, 1) = ")" Then
          open_next = True
       Else
          open_next = False
       End If
End Function

