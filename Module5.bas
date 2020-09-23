Attribute VB_Name = "Module5"

Public Sub Engine(char As String, length As Byte)

   ReDim ary(length)
   Dim depth As Byte
   Dim result As String
   Dim flag As Boolean
   flag = False
   depth = 1

   While Not (flag)
         k = DoEvents()
         ary(depth) = ary(depth) + 1

         If depth = length Then
            result = Left(result, length - 1) + Mid(char, ary(depth), 1)
         Else
            result = result + Mid(char, ary(depth), 1)
         End If

         If ary(depth) <> Len(char) + 1 Then
            If depth <> length Then
               depth = depth + 1
            Else
               '****************************************
               Dim temp As String
               Dim i As Integer
               Dim j As Integer
               Dim l As Integer
               Dim expr_result As String
               Dim p As Byte
               expr_result = ""
               With Form1
                    For i = 0 To .lstexpr.ListCount - 1
                        temp = .lstexpr.List(i)
                        Do
                            For j = 0 To .lstvar.ListCount - 1
                                If extract_variable(temp, p) = .lstvar.List(j) Then
                                   temp = del(temp, p, Len(.lstvar.List(j)))
                                   temp = insert(Mid(result, j + 1, 1), temp, p)
                                End If
                             Next j
                        Loop While extract_variable(temp, p) <> ""
                        expr_result = expr_result + evaluate(temp)
                    Next i
                    Call morfi(result, expr_result)
                    .Prog.Value = .Prog.Value + 1
                    .Label1.Caption = Str(.Prog.Value) + "/" + Str(.Prog.Max)
               End With
               
               '****************************************
            End If
         Else
            If depth = 1 Then
               flag = True
            Else
               ary(depth) = 0
               depth = depth - 1
               result = Left(result, depth - 1)
            End If
         End If
   Wend

  
End Sub

Public Function extract_variable(expr As String, pos As Byte) As String
       Dim i As Byte
       Dim temp As String
       Dim flag As Boolean
       flag = True
       For i = 1 To Len(expr)
           If character(Mid(expr, i, 1)) Then
              If flag = True Then
                 pos = i: flag = False
              End If
              temp = temp + Mid(expr, i, 1)
           Else
              If Len(temp) <> 0 Then
                 Exit For
              End If
           End If
       Next i
       extract_variable = temp
End Function


Public Sub morfi(variables As String, exprs As String)
       Dim i As Integer
       Dim temp As String
       temp = Space(5)
       With Form1
            For i = 1 To Len(variables)
                temp = temp + Space(Len(.lstvar.List(i - 1)) - (Len(.lstvar.List(i - 1)) \ 2) - 1)
                temp = temp + Mid(variables, i, 1)
                If Len(.lstvar.List(i - 1)) Mod 2 = 0 Then
                   temp = temp + Space(Len(.lstvar.List(i - 1)) - (Len(.lstvar.List(i - 1)) \ 2))
                Else
                   temp = temp + Space(Len(.lstvar.List(i - 1)) - (Len(.lstvar.List(i - 1)) \ 2) - 1)
                End If
                temp = temp + Space(2)
            Next i
            temp = temp + Space(5)
            '*******************************************
            For i = 1 To Len(exprs)
                temp = temp + Space(Len(.lstexpr.List(i - 1)) - (Len(.lstexpr.List(i - 1)) \ 2) - 1)
                temp = temp + Mid(exprs, i, 1)
                If Len(.lstexpr.List(i - 1)) Mod 2 = 0 Then
                   temp = temp + Space(Len(.lstexpr.List(i - 1)) - (Len(.lstexpr.List(i - 1)) \ 2))
                Else
                   temp = temp + Space(Len(.lstexpr.List(i - 1)) - (Len(.lstexpr.List(i - 1)) \ 2) - 1)
                End If
                temp = temp + Space(2)
            Next i
            Print #1, temp
       End With
End Sub

Public Sub maska()
       Dim temp As String
       Dim i As Integer
       temp = Space(5)
       With Form1
            For i = 0 To .lstvar.ListCount - 1
                temp = temp + .lstvar.List(i)
                temp = temp + Space(2)
            Next i
            temp = temp + Space(5)
            For i = 0 To .lstexpr.ListCount - 1
                temp = temp + .lstexpr.List(i)
                temp = temp + Space(2)
            Next i
       End With
       Print #1, temp
       Print #1, String(Len(temp), "-")
End Sub
