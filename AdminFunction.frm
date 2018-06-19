VERSION 5.00
Begin VB.Form AdminFunctions 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aptitude Test And Career Guidance"
   ClientHeight    =   7560
   ClientLeft      =   7050
   ClientTop       =   480
   ClientWidth     =   4905
   DrawMode        =   5  'Not Copy Pen
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdViewQuestions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Questions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   2004
   End
   Begin VB.CommandButton cmdViewTeacherDetails 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Teacher Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   2004
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   732
   End
   Begin VB.CommandButton cmdDeleteQuestion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete Question"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   2004
   End
   Begin VB.CommandButton cmdAddQuestion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Question"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   2004
   End
   Begin VB.CommandButton cmdRemoveTeacher 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete Teacher"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1560
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   2004
   End
   Begin VB.CommandButton cmdAddTeacher 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Teacher"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   2004
   End
   Begin VB.Label Label1 
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
      Height          =   975
      Left            =   600
      TabIndex        =   8
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label w 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome Admin!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
End
Attribute VB_Name = "AdminFunctions"
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

Private Sub cmdViewQuestions_Click()
Unload Me
DisplayQuestion.Show
End Sub

Private Sub cmdViewTeacherDetails_Click()
Unload Me
DisplayTeacherDetails.Show
End Sub
