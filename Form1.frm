VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Truth Table Generator by Palladinos Nick  e-mail: codikas@x-treme.gr"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   2160
      ScaleHeight     =   2115
      ScaleWidth      =   4635
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   4695
      Begin MSComctlLib.ProgressBar Prog 
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Enter Expression"
      Height          =   855
      Left            =   2160
      TabIndex        =   5
      Top             =   720
      Width           =   4335
      Begin VB.CommandButton Command2 
         Caption         =   "Add"
         Height          =   495
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Text            =   "~a+(bad*~(1+a)*~0+good)"
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Expressions"
      Height          =   3615
      Left            =   4800
      TabIndex        =   2
      Top             =   1800
      Width           =   3855
      Begin VB.ListBox lstexpr 
         Height          =   2985
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Variables"
      Height          =   3615
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   3855
      Begin VB.ListBox lstvar 
         Height          =   2985
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   5640
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
        If lstvar.ListCount <> 0 Then
           MsgBox "Saving in " + App.Path + "\result.txt"
           Open App.Path + "\result.txt" For Output As #1
           Call maska
           '**************************************
           Command1.Enabled = False
           Command2.Enabled = False
           Picture1.Visible = True
           Prog.Min = 0
           Prog.Value = 0
           Prog.Max = 2 ^ lstvar.ListCount
           Call Engine("01", lstvar.ListCount)
           Picture1.Visible = False
           Command1.Enabled = True
           Command2.Enabled = True
           '***************************************
           Close #1
        Else
           MsgBox "No variable found"
        End If
End Sub

Public Sub add_expr(expr As String)
       Dim i As Byte
       Dim flag As Boolean
       flag = True
       For i = 0 To lstexpr.ListCount
           If lstexpr.List(i) = expr Then
              flag = False: Exit For
           End If
       Next i
       If flag = True Then
          lstexpr.AddItem expr
       End If
End Sub
Public Sub extract_variables(expr As String)
       Dim i As Byte
       Dim j As Byte
       Dim flag  As Boolean
       Dim temp As String
       For i = 1 To Len(expr)
           If character(Mid(expr, i, 1)) Then
              Do
                   temp = temp + Mid(expr, i, 1)
                   i = i + 1
              Loop While character(Mid(expr, i, 1))
              flag = True
              For j = 0 To lstvar.ListCount
                  If lstvar.List(j) = temp Then
                     flag = False
                  End If
              Next j
              If flag = True Then
                 lstvar.AddItem temp
              End If
              temp = ""
           End If
       Next i
End Sub


Private Sub Command2_Click()
        Dim temp As String
        temp = alltrim(Text1.Text)
        If valid_expr(temp) Then
           Call extract_variables(Text1.Text)
           Call add_expr(Text1.Text)
        End If
End Sub
