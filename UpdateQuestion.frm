VERSION 5.00
Begin VB.Form DeleteQuestionPage 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aptitude Test And Career Guidance"
   ClientHeight    =   7560
   ClientLeft      =   7050
   ClientTop       =   480
   ClientWidth     =   4800
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   948
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
      Height          =   504
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1170
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
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
      TabIndex        =   2
      Top             =   2280
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
      TabIndex        =   4
      Top             =   2280
      Width           =   684
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
      TabIndex        =   7
      Top             =   0
      Width           =   732
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      Height          =   372
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtDelete 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label labMessage 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label Label2 
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
      Left            =   840
      TabIndex        =   9
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter no.  of  question which is to be deleted"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Width           =   2535
   End
End
Attribute VB_Name = "DeleteQuestionPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As New Connection
Dim R As New Recordset
Dim S As String

Private Sub cmdAddQuestion_Click()
Unload Me
QuestionScreen.Show
End Sub

Private Sub cmdAddTeacher_Click()
Unload Me
AddTeacherPage.Show
End Sub

Private Sub cmdDelete_Click()

Dim id As Integer
id = Val(txtDelete.Text)
C.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\Questions.mdb;Persist Security Info=False"
S = "Select * from QuestionTable Where(qno = " & id & ")"
R.Open S, C, adOpenDynamic, adLockOptimistic
If Not R.EOF And Not R.BOF Then
    R.Close
    S = "Delete from QuestionTable where (qno=" & id & ")"
    R.Open S, C, adOpenDynamic, adLockOptimistic
    labMessage.Visible = True
    labMessage.ForeColor = vbRed
    labMessage.BackColor = vbWhite
    labMessage.BackStyle = 1
    labMessage.Caption = "Record is deleted."
Else
    labMessage.Visible = True
    labMessage.Caption = "Record is not present."
    labMessage.ForeColor = vbRed
    labMessage.BackColor = vbWhite
    labMessage.BackStyle = 1
End If
C.Close
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


Private Sub txtDelete_KeyPress(KeyAscii As Integer)
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
