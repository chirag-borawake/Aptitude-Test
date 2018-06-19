VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form QuestionScreen 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aptitude Test And Career Guidance"
   ClientHeight    =   7560
   ClientLeft      =   7050
   ClientTop       =   480
   ClientWidth     =   4815
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optTechnical 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Technical"
      ForeColor       =   &H8000000F&
      Height          =   192
      Left            =   2040
      TabIndex        =   8
      Top             =   6000
      Width           =   1095
   End
   Begin VB.OptionButton optFinance 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Finance"
      ForeColor       =   &H8000000F&
      Height          =   192
      Left            =   2040
      TabIndex        =   10
      Top             =   6360
      Width           =   972
   End
   Begin VB.OptionButton optArts 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Arts"
      ForeColor       =   &H8000000F&
      Height          =   192
      Left            =   3240
      TabIndex        =   11
      Top             =   6360
      Width           =   972
   End
   Begin VB.OptionButton optMedical 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Medical"
      ForeColor       =   &H8000000F&
      Height          =   192
      Left            =   3240
      TabIndex        =   9
      Top             =   6000
      Width           =   972
   End
   Begin VB.CommandButton cmdViewQuestions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Questions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1320
      Width           =   804
   End
   Begin VB.TextBox txtOpt4 
      BackColor       =   &H00FFFFFF&
      DataField       =   "opt4"
      DataSource      =   "QuestionADODC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1920
      TabIndex        =   6
      Text            =   " "
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox txtAns 
      BackColor       =   &H00FFFFFF&
      DataField       =   "ans"
      DataSource      =   "QuestionADODC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1920
      TabIndex        =   7
      Text            =   " "
      Top             =   5520
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc QuestionADODC 
      Height          =   312
      Left            =   1440
      Top             =   0
      Visible         =   0   'False
      Width           =   1692
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      Height          =   372
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   972
   End
   Begin VB.CommandButton cmdViewTeacherDetails 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Teacher Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1320
      Width           =   1170
   End
   Begin VB.CommandButton cmdAddTeacher 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Teacher"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   804
   End
   Begin VB.CommandButton cmdRemoveTeacher 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete Teacher"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2160
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1320
      Width           =   804
   End
   Begin VB.CommandButton cmdDeleteQuestion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete Question"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1320
      Width           =   804
   End
   Begin VB.CommandButton cmdLogOut 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Log out"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   0
      Width           =   732
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New"
      Height          =   372
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   1092
   End
   Begin VB.TextBox txtOpt3 
      BackColor       =   &H00FFFFFF&
      DataField       =   "opt3"
      DataSource      =   "QuestionADODC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   " "
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox txtOpt2 
      BackColor       =   &H00FFFFFF&
      DataField       =   "opt2"
      DataSource      =   "QuestionADODC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   " "
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox txtOpt1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "opt1"
      DataSource      =   "QuestionADODC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1920
      TabIndex        =   3
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox txtQuestion 
      BackColor       =   &H00FFFFFF&
      DataField       =   "question"
      DataSource      =   "QuestionADODC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   888
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtQno 
      BackColor       =   &H00FFFFFF&
      DataField       =   "qno"
      DataSource      =   "QuestionADODC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1920
      TabIndex        =   1
      Text            =   " "
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Categary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   28
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label labMessage 
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
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   27
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   26
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Aptitude Test And Career Guidance"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Left            =   720
      TabIndex        =   25
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Option 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   24
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Option 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   23
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Option 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   22
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Option 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Question"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   20
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Question no."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "QuestionScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    

Private Sub cmdAddQuestion_Click()
Unload Me
End Sub

Private Sub cmdAddTeacher_Click()
Unload Me
AddTeacherPage.Show
End Sub

Private Sub cmdDeleteQuestion_Click()
Unload Me
DeleteQuestionPage.Show
End Sub

Private Sub cmdLogOut_Click()
Unload Me
mainpage.Show
End Sub

Private Sub cmdNew_Click()
QuestionADODC.Recordset.AddNew
txtQno.Text = ""
txtQuestion.Text = ""
txtOpt1.Text = ""
txtOpt2.Text = ""
txtOpt3.Text = ""
txtOpt4.Text = ""
optArts.Value = False
optTechnical.Value = False
optMedical.Value = False
optFinance.Value = False
labMessage.Visible = False
txtQno.SetFocus
End Sub

Private Sub cmdRemoveTeacher_Click()
Unload Me
DeleteTeacher.Show
End Sub


Private Sub cmdSave_Click()
QuestionADODC.Recordset.Fields(0) = txtQno.Text
QuestionADODC.Recordset.Fields(1) = txtQuestion.Text
QuestionADODC.Recordset.Fields(2) = txtOpt1.Text
QuestionADODC.Recordset.Fields(3) = txtOpt2.Text
QuestionADODC.Recordset.Fields(4) = txtOpt3.Text
QuestionADODC.Recordset.Fields(5) = txtOpt4.Text
QuestionADODC.Recordset.Fields(6) = txtAns.Text
If optMedical.Value = True Then
    QuestionADODC.Recordset.Fields(7) = "Medical"
ElseIf optTechnical.Value = True Then
    QuestionADODC.Recordset.Fields(7) = "Technical"
ElseIf optFinance.Value = True Then
    QuestionADODC.Recordset.Fields(7) = "Finance"
ElseIf optArts.Value = True Then
    QuestionADODC.Recordset.Fields(7) = "Arts"
End If
QuestionADODC.Recordset.MoveNext
QuestionADODC.Refresh
If Not QuestionADODC.Recordset.EOF And Not QuestionADODC.Recordset.BOF Then
labMessage.Visible = True
labMessage.Caption = " Question Added Successfully!"
labMessage.ForeColor = vbRed
labMessage.BackColor = vbWhite
labMessage.BackStyle = 1
End If
End Sub


Private Sub cmdViewQuestions_Click()
Unload Me
DisplayQuestion.Show
End Sub

Private Sub cmdViewTeacherDetails_Click()
Unload Me
DisplayTeacherDetails.Show
End Sub


Private Sub txtQno_KeyPress(KeyAscii As Integer)
KeyAscii = integerOnly(KeyAscii)
End Sub
Private Function integerOnly(intKey As Integer) As Integer
  Dim intReturn As Integer
  intReturn = intKey
  Select Case intKey
    Case vbKeyBack        'allow backspace
    Case vbKey0 To vbKey9 'allow numbers 0 to 9
    Case Else             'block all other input
      intReturn = 0
  End Select
  integerOnly = intReturn
End Function

Private Sub txtCatagory_KeyPress(KeyAscii As Integer)
KeyAscii = alphabetOnly(KeyAscii)
End Sub
Private Function alphabetOnly(intKey As Integer) As Integer
  Dim intReturn As Integer
  intReturn = intKey
  Select Case intKey
    Case vbKeyBack        'allow backspace
    Case vbKeyA To vbKeyZ 'allow numbers A to Z
    Case 97 To 122        'allow a to z
    Case Else             'block all other input
      intReturn = 0
  End Select
  alphabetOnly = intReturn
End Function
Private Sub txtQuestion_KeyPress(KeyAscii As Integer)
KeyAscii = integerAndAlphabetsSymbols(KeyAscii)
End Sub

Private Sub txtOpt1_KeyPress(KeyAscii As Integer)
 KeyAscii = integerAndAlphabetsSymbols(KeyAscii)
End Sub

Private Sub txtOpt2_KeyPress(KeyAscii As Integer)
 KeyAscii = integerAndAlphabetsSymbols(KeyAscii)
End Sub

Private Sub txtOpt3_KeyPress(KeyAscii As Integer)
 KeyAscii = integerAndAlphabetsSymbols(KeyAscii)
End Sub


Private Sub txtOpt4_KeyPress(KeyAscii As Integer)
 KeyAscii = integerAndAlphabetsSymbols(KeyAscii)
End Sub

Private Sub txtAns_KeyPress(KeyAscii As Integer)
 KeyAscii = integerAndAlphabetsSymbols(KeyAscii)
End Sub
Private Function integerAndAlphabetsSymbols(intKey As Integer) As Integer
  Dim intReturn As Integer
  intReturn = intKey
  Select Case intKey
    Case vbKeyBack        'allow backspace
    Case 32 To 122
    Case Else             'block all other input
      intReturn = 0
  End Select
  integerAndAlphabetsSymbols = intReturn
End Function
