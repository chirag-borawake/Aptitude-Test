VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form testForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aptitude Test And Career Guidance"
   ClientHeight    =   7560
   ClientLeft      =   7050
   ClientTop       =   480
   ClientWidth     =   4905
   LinkTopic       =   "Test"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   0
      Top             =   6480
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1800
      TabIndex        =   19
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   6840
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6720
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   5360
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   3960
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   3240
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc adoGetQuestion 
      Height          =   270
      Left            =   0
      Top             =   7320
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   476
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\Questions.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\Questions.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "QuestionTable"
      Caption         =   "Question"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label labTimeUsed 
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labMinute 
      BackStyle       =   0  'Transparent
      Caption         =   "49"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3600
      TabIndex        =   18
      Top             =   240
      Width           =   375
   End
   Begin VB.Label labSeconds 
      BackStyle       =   0  'Transparent
      Caption         =   "59"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4065
      TabIndex        =   17
      Top             =   240
      Width           =   360
   End
   Begin VB.Label labDisplay 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4065
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labSymbol 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3975
      TabIndex        =   15
      Top             =   270
      Width           =   135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Question"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      DataField       =   "qno"
      DataSource      =   "adoGetQuestion"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " /30    Time Left"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lebMessage 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "A)"
      DataField       =   "opt1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "B)"
      DataField       =   "opt2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "C)"
      DataField       =   "opt3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "D)"
      DataField       =   "opt4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label LabQno 
      BackStyle       =   0  'Transparent
      DataField       =   "qno"
      DataSource      =   "adoGetQuestion"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   255
   End
   Begin VB.Label labQuestion 
      BackStyle       =   0  'Transparent
      DataField       =   "question"
      DataSource      =   "adoGetQuestion"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1695
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   3735
   End
End
Attribute VB_Name = "testForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim starttime As Date
Dim starttime1 As Date
Dim C As New Connection
Dim R  As New Recordset
Dim sql As String
Dim C1 As New Connection
Dim R1  As New Recordset
Dim S As String
Dim S1 As String
Dim technical  As Integer
Dim medical As Integer
Dim finance As Integer
Dim arts As Integer
Dim answer As Variant
Dim CorrectAnswer As Integer
Dim WrongAnswer As Integer

Private Sub cmdSubmit_Click()
If Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False Then
    lebMessage.Visible = True
    lebMessage.Caption = "You have to select option"
    lebMessage.ForeColor = vbRed
    lebMessage.BackColor = vbWhite
    lebMessage.BackStyle = 1
Else
    If Not adoGetQuestion.Recordset.EOF Then
        lebMessage.Visible = False
        If adoGetQuestion.Recordset.Fields(7).Value = "Technical" Then
            If Option1.Value = True Then
                answer = Option1.Caption
            ElseIf Option2.Value = True Then
                answer = Option2.Caption
            ElseIf Option3.Value = True Then
                answer = Option3.Caption
            Else
                answer = Option4.Caption
            End If
            If adoGetQuestion.Recordset.Fields(6) = answer Then
                CorrectAnswer = CorrectAnswer + 1
                technical = technical + 1
            Else
                WrongAnswer = WrongAnswer + 1
            End If
        ElseIf adoGetQuestion.Recordset.Fields(7).Value = "Medical" Then
            If Option1.Value = True Then
                answer = Option1.Caption
            ElseIf Option2.Value = True Then
                answer = Option2.Caption
            ElseIf Option3.Value = True Then
                answer = Option3.Caption
            Else
                answer = Option4.Caption
            End If
            If adoGetQuestion.Recordset.Fields(6) = answer Then
                CorrectAnswer = CorrectAnswer + 1
                medical = medical + 1
            Else
                WrongAnswer = WrongAnswer + 1
            End If
        ElseIf adoGetQuestion.Recordset.Fields(7).Value = "Finance" Then
            If Option1.Value = True Then
                answer = Option1.Caption
            ElseIf Option2.Value = True Then
                answer = Option2.Caption
            ElseIf Option3.Value = True Then
                answer = Option3.Caption
            Else
                answer = Option4.Caption
            End If
            If adoGetQuestion.Recordset.Fields(6) = answer Then
                CorrectAnswer = CorrectAnswer + 1
                finance = finance + 1
            Else
                WrongAnswer = WrongAnswer + 1
            End If
        ElseIf adoGetQuestion.Recordset.Fields(7).Value = "Arts" Then
            If Option1.Value = True Then
                answer = Option1.Caption
            ElseIf Option2.Value = True Then
                answer = Option2.Caption
            ElseIf Option3.Value = True Then
                answer = Option3.Caption
            Else
                answer = Option4.Caption
            End If
            If adoGetQuestion.Recordset.Fields(6) = answer Then
                CorrectAnswer = CorrectAnswer + 1
                arts = arts + 1
            Else
                WrongAnswer = WrongAnswer + 1
            End If
            
        End If
        If Not adoGetQuestion.Recordset.EOF Then
            adoGetQuestion.Recordset.MoveNext
        End If
        If Not adoGetQuestion.Recordset.EOF Then
            Option1.Value = False
            Option2.Value = False
            Option3.Value = False
            Option4.Value = False
            Option1.Caption = adoGetQuestion.Recordset.Fields(2).Value
            Option2.Caption = adoGetQuestion.Recordset.Fields(3).Value
            Option3.Caption = adoGetQuestion.Recordset.Fields(4).Value
            Option4.Caption = adoGetQuestion.Recordset.Fields(5).Value
        End If
    End If
    If adoGetQuestion.Recordset.EOF Then
        adoGetQuestion.Recordset.MovePrevious
        cmdSubmit.Enabled = False
        C.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\Student.mdb;Persist Security Info=False"
        sql = "Select * From  StudentTable"
        R.Open sql, C, adOpenDynamic, adLockOptimistic
        If Not R.BOF And Not R.EOF Then
            R.MoveFirst
            While Not R.Fields(8) = txtUsername.Text
                R.MoveNext
            Wend
            C1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\StudentResult.mdb;Persist Security Info=False"
            S1 = "Insert Into StudentResultTable Values(" & R.Fields(0).Value & ",'" & R.Fields(1).Value & "'," & technical & "," & medical & ", " & arts & "," & finance & "," & CorrectAnswer & "," & WrongAnswer & ")"
            R1.Open S1, C1, adOpenDynamic, adLockOptimistic
            ResultPage.labCorrectAnswer = CorrectAnswer
            ResultPage.labWrongAnswer = WrongAnswer
            chart.txtTechnical = technical
            chart.txtMedical = medical
            chart.txtArts = arts
            chart.txtFinance = finance
            ResultPage.labTimeUsed = Me.labTimeUsed
            Unload Me
            ResultPage.Show
        End If
        R.Close
        C.Close
       End If
End If
End Sub
Private Sub Timer2_Timer()
labTimeUsed.Caption = Format$(Now - starttime1, "hh:mm:ss")
End Sub
Private Sub Form_Load()
Timer1.Enabled = True
starttime1 = Now
Timer2.Enabled = True
technical = 0
medical = 0
finance = 0
arts = 0
WrongAnswer = 0
CorrectAnswer = 0
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option1.Caption = adoGetQuestion.Recordset.Fields(2).Value
Option2.Caption = adoGetQuestion.Recordset.Fields(3).Value
Option3.Caption = adoGetQuestion.Recordset.Fields(4).Value
Option4.Caption = adoGetQuestion.Recordset.Fields(5).Value
End Sub


Private Sub Timer1_Timer()
labSeconds.Caption = labSeconds.Caption - 1
If labSeconds.Caption = 0 And labMinute.Caption = 0 Then
    MsgBox "Your time is over."
     C.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\Student.mdb;Persist Security Info=False"
        sql = "Select * From  StudentTable"
        R.Open sql, C, adOpenDynamic, adLockPessimistic
        If Not R.BOF And Not R.EOF Then
            R.MoveFirst
            While Not R.Fields(8) = txtUsername.Text
                R.MoveNext
            Wend
            C1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\StudentResult.mdb;Persist Security Info=False"
            S = "Insert Into StudentResultTable Values(" & R.Fields(0).Value & ",'" & R.Fields(1).Value & "'," & technical & "," & medical & ", " & arts & "," & finance & "," & CorrectAnswer & "," & WrongAnswer & ")"
            R1.Open S, C1, adOpenDynamic
            Timer1.Enabled = False
            ResultPage.labCorrectAnswer = CorrectAnswer
            ResultPage.labWrongAnswer = WrongAnswer
            ResultPage.labTimeUsed = Me.labTimeUsed
            Unload Me
            ResultPage.Show
        End If
        R.Close
        C.Close
End If
If labSeconds.Caption = -1 Then
labMinute.Caption = labMinute.Caption - 1
labSeconds.Caption = 59
End If
If labSeconds.Caption > 10 Then
    labDisplay.Visible = False
    labSeconds.Left = 4070
End If
If labSeconds.Caption < 10 Then
    labDisplay.Visible = True
    labSeconds.Left = 4230
End If
End Sub


