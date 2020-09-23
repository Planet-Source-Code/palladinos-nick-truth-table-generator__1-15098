Attribute VB_Name = "Module1"
Public Function alltrim(expr As String) As String
       Dim temp As String
       Dim pos As Byte
       temp = expr
       temp = Trim(temp)
       pos = InStr(1, temp, " ")
       Do While pos <> 0
          temp = del(temp, pos, 1)
          pos = InStr(1, temp, " ")
       Loop
       alltrim = temp
End Function

Public Function del(s As String, start As Byte, length As Byte)
       del = Left(s, start - 1) + Right(s, Len(s) - (start + (length - 1)))
End Function
Public Function insert(input_string As String, s As String, start As Byte)
       insert = Left(s, start - 1) + input_string + Right(s, (Len(s) - start) + 1)
End Function


Public Function valid_expr(expr As String) As Boolean
       Dim i As Byte
       valid_expr = True
       For i = 1 To Len(expr)
           '***************************************
           If character(Mid(expr, i, 1)) Then
              If Not (character_prev(expr, i) And character_next(expr, i)) Then
                 valid_expr = False: Exit Function
              End If
           End If
           '******************************************
           If zeroone_validation(Mid(expr, i, 1)) Then
              If Not (zeroone_prev(expr, i) And zeroone_next(expr, i)) Then
                 valid_expr = False: Exit Function
              End If
           End If
           '*********************************************
           If Mid(expr, i, 1) = "*" Or Mid(expr, i, 1) = "+" Then
              If Not (andor_prev(expr, i) And andor_next(expr, i)) Then
                 valid_expr = False: Exit Function
              End If
           End If
           '**************************************************
           If Mid(expr, i, 1) = "~" Then
              If Not (not_prev(expr, i) And not_next(expr, i)) Then
                 valid_expr = False: Exit Function
              End If
           End If
           '***************************************************
           If Mid(expr, i, 1) = "(" Then
              If open_next(expr, i) Then
                 valid_expr = False: Exit Function
              End If
           End If
           '****************************************************
       Next i
       If Not (valid_parenthesis(expr)) Then
          valid_expr = False
       End If
End Function
