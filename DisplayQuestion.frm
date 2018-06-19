VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form DisplayQuestion 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aptitude Test And Career Guidance"
   ClientHeight    =   7560
   ClientLeft      =   7050
   ClientTop       =   480
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCatagory 
      BackColor       =   &H00FFFFFF&
      DataField       =   "cat"
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   7200
      Width           =   2295
   End
   Begin VB.TextBox txtQno 
      BackColor       =   &H00FFFFFF&
      DataField       =   "qno"
      DataSource      =   "QuestionADODC"
      Height          =   288
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   " "
      Top             =   2760
      Width           =   2295
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
      Height          =   1365
      Left            =   1800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3240
      Width           =   2295
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4800
      Width           =   2295
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
      Height          =   288
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   " "
      Top             =   5280
      Width           =   2295
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
      Height          =   288
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   " "
      Top             =   5760
      Width           =   2295
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
      TabIndex        =   6
      Top             =   0
      Width           =   732
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   1320
      Width           =   804
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
      TabIndex        =   0
      Top             =   1320
      Width           =   804
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
      TabIndex        =   1
      Top             =   1320
      Width           =   1170
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   " "
      Top             =   6720
      Width           =   2295
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   " "
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton cmdAddQuestion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Question"
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   804
   End
   Begin MSAdodcLib.Adodc QuestionADODC 
      Height          =   315
      Left            =   1440
      Top             =   1920
      Width           =   2055
      _ExtentX        =   3625
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
      BackColor       =   16777215
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
      Left            =   720
      TabIndex        =   22
      Top             =   7200
      Width           =   855
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
      Height          =   255
      Left            =   1560
      TabIndex        =   21
      Top             =   2400
      Width           =   1695
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
      Left            =   480
      TabIndex        =   20
      Top             =   2760
      Width           =   1215
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
      Left            =   720
      TabIndex        =   19
      Top             =   3240
      Width           =   975
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
      Left            =   720
      TabIndex        =   18
      Top             =   4800
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
      Left            =   720
      TabIndex        =   17
      Top             =   5280
      Width           =   855
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
      Left            =   720
      TabIndex        =   16
      Top             =   5760
      Width           =   855
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
      Left            =   720
      TabIndex        =   15
      Top             =   6240
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
      TabIndex        =   14
      Top             =   240
      Width           =   3375
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
      TabIndex        =   13
      Top             =   6720
      Width           =   735
   End
End
Attribute VB_Name = "DisplayQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddQuestion_Click()
Unload Me
QuestionScreen.Show
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

Private Sub cmdRemoveTeacher_Click()
Unload Me
DeleteTeacher.Show

End Sub

Private Sub cmdViewTeacherDetails_Click()
Unload Me
DisplayTeacherDetails.Show
End Sub

Private Sub QuestionADODC_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If Not QuestionADODC.Recordset.BOF And Not QuestionADODC.Recordset.EOF Then
    labMessage.Visible = False
Else
    labMessage.Visible = True
    labMessage.Caption = " No more records."
    labMessage.ForeColor = vbRed
    labMessage.BackColor = vbWhite
    labMessage.BackStyle = 1
End If
End Sub


