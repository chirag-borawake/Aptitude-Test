VERSION 5.00
Begin VB.Form InstructionPage 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aptitude Test And Career Guidance"
   ClientHeight    =   7560
   ClientLeft      =   7050
   ClientTop       =   480
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUsername 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   -120
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdBegin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Begin"
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label labMessage2 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Wel - Come"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label labMessage3 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "        When you are ready to start, click 'Begin' to start the test."
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   1215
      Left            =   360
      TabIndex        =   4
      Top             =   4200
      Width           =   4215
   End
   Begin VB.Label labMessage 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "          This test consists of 40 mixed aptitude  questions,  you   have   50   minutes   to complete the test."
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   1935
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label labName 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   855
      Width           =   1215
   End
End
Attribute VB_Name = "InstructionPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As New Connection
Dim R  As New Recordset
Dim sql As String
Private Sub cmdBegin_Click()
Unload Me
testForm.Show
End Sub

Private Sub txtUsername_Change()
C.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\Student.mdb;Persist Security Info=False"
sql = "Select * From  StudentTable where username ='" & txtUsername.Text & "'"
R.Open sql, C, adOpenDynamic, adLockOptimistic
If Not R.BOF And Not R.EOF Then
    labName.Caption = R.Fields(1)
End If
R.Close
C.Close
End Sub
