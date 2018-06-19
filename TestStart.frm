VERSION 5.00
Begin VB.Form TestStartupScreen 
   Appearance      =   0  'Flat
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aptitude Test And Career Guidance"
   ClientHeight    =   7560
   ClientLeft      =   7050
   ClientTop       =   480
   ClientWidth     =   4905
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSignIn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sign in"
      Height          =   372
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1212
   End
   Begin VB.TextBox txtPassword 
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
      Height          =   372
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4680
      Width           =   2052
   End
   Begin VB.TextBox txtUsername 
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
      Height          =   372
      Left            =   1320
      TabIndex        =   0
      Top             =   3600
      Width           =   2052
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Student Log in"
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
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lebMessage 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   840
      TabIndex        =   6
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User name "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
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
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
End
Attribute VB_Name = "TestStartupScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As New Connection
Dim R As New Recordset
Dim S, sql As String

Private Sub cmdSignIn_Click()
C.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\Student.mdb;Persist Security Info=False"
sql = "Select * From  StudentTable where username='" & Trim(txtUsername.Text) & "' and password='" & Trim(txtPassword.Text) & "'"
R.Open sql, C, adOpenDynamic, adLockOptimistic
If Not R.BOF And Not R.EOF Then
    lebMessage.Visible = False
    Me.Hide
    InstructionPage.Show
    InstructionPage.txtUsername = Me.txtUsername.Text
    testForm.txtUsername = Me.txtUsername.Text
Else
    lebMessage.Caption = "The username or password you entered is incorrect."
    lebMessage.ForeColor = vbRed
    lebMessage.BackColor = vbWhite
    lebMessage.BackStyle = 1
End If
R.Close
C.Close
End Sub



