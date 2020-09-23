Attribute VB_Name = "Module3"
Public Function evaluate(ByVal expr As String) As String
       Dim flag As Boolean
       Dim pos As Byte
       Dim pos1 As Byte
       Dim temp As String
       flag = True
       pos = Len(expr)
       Do
            pos = pos - 1
            If pos > 0 Then
               If Mid(expr, pos, 1) = "(" Then
                  pos1 = pos + 1
                  Do
                       temp = temp + Mid(expr, pos1, 1)
                       pos1 = pos1 + 1
                  Loop While Mid(expr, pos1, 1) <> ")"
                  expr = del(expr, pos, Len(temp) + 2)
                  temp = parse(temp)
                  expr = insert(temp, expr, pos)
                  pos = Len(expr)
                  temp = ""
               End If
            Else
               flag = False
            End If
       Loop While flag
       expr = parse(expr)
       evaluate = expr
End Function
Public Function parse(ByVal expr As String) As String
       Dim flag As Boolean
       Dim pos As Byte
       Dim pos1 As Byte
       Dim stage As Byte
       Dim temp As String
       stage = 1
       flag = True
       If Len(expr) = 1 Then
          parse = expr: Exit Function
       End If
       Do
            '*******************************************
            If stage = 1 Then
               pos = InStr(1, expr, "~")
               If pos <> 0 Then
                  temp = not_parse(Mid(expr, pos + 1, 1))
                  expr = del(expr, pos, 2)
                  expr = insert(temp, expr, pos)
               Else
                  stage = stage + 1
               End If
            End If
            '******************************************
            If stage = 2 Then
               pos = 0
               Do
                    pos = pos + 1
                    If Mid(expr, pos, 1) = "*" Then
                       temp = and_parse(Mid(expr, pos - 1, 1), Mid(expr, pos + 1, 1))
                       expr = del(expr, pos - 1, 3)
                       expr = insert(temp, expr, pos - 1)
                       pos = pos - 1
                    End If
                    If Mid(expr, pos, 1) = "+" Then
                       temp = or_parse(Mid(expr, pos - 1, 1), Mid(expr, pos + 1, 1))
                       expr = del(expr, pos - 1, 3)
                       expr = insert(temp, expr, pos - 1)
                       pos = pos - 1
                    End If
               Loop While pos <= Len(expr)
               flag = False
            End If
       Loop While flag
       parse = expr
End Function
Public Function or_parse(expr As String, expr1 As String) As String
       If expr = "0" And expr1 = "0" Then
          or_parse = "0"
       Else
          or_parse = "1"
       End If
End Function
Public Function and_parse(expr As String, expr1 As String) As String
       If expr = "1" And expr1 = "1" Then
          and_parse = "1"
       Else
          and_parse = "0"
       End If
End Function
Public Function not_parse(expr As String) As String
       If expr = "1" Then
          not_parse = "0"
       Else
          not_parse = "1"
       End If
End Function
