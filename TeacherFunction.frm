VERSION 5.00
Begin VB.Form TeacherFunctions 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aptitude Test And Career Guidance"
   ClientHeight    =   7560
   ClientLeft      =   7050
   ClientTop       =   480
   ClientWidth     =   4905
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   150
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   732
   End
   Begin VB.CommandButton cmdGetResult 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get Result"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   2115
   End
   Begin VB.CommandButton cmdViewDetails 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Student Details"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   2115
   End
   Begin VB.CommandButton cmdGetTest 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get Test"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   2115
   End
   Begin VB.CommandButton cmdRemoveStudent 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete Student"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   2115
   End
   Begin VB.CommandButton cmdAddStudent 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Student"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   2115
   End
   Begin VB.Label labName 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
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
      Left            =   1560
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label w 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
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
      Height          =   372
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   972
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
      Height          =   975
      Left            =   480
      TabIndex        =   6
      Top             =   600
      Width           =   3975
   End
End
Attribute VB_Name = "TeacherFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As New Connection
Dim R As New Recordset
Dim S, sql As String

Private Sub cmdAddStudent_Click()
Unload Me
GetStudentDetails.Show
End Sub

Private Sub cmdGetResult_Click()
Unload Me
ResultPage.Show
End Sub

Private Sub cmdGetTest_Click()
Unload Me
TestStartupScreen.Show
End Sub

Private Sub cmdLogOut_Click()
Unload Me
Unload mainpage
mainpage.Show
End Sub

Private Sub cmdRemoveStudent_Click()
Unload Me
DeleteStudent.Show
End Sub


Private Sub cmdViewDetails_Click()
Unload Me
StudentDetailsPage.Show
End Sub


Private Sub txtUsername_Change()
C.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\Teacher.mdb;Persist Security Info=False"
sql = "Select * From  TeacherTable where username ='" & txtUsername.Text & "'"
R.Open sql, C, adOpenDynamic, adLockOptimistic
If Not R.BOF And Not R.EOF Then
    labName.Caption = R.Fields(1)
End If
R.Close
C.Close
End Sub
