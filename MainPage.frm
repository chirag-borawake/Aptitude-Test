VERSION 5.00
Begin VB.Form MainPage 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aptitude Test And Career Guidance"
   ClientHeight    =   7560
   ClientLeft      =   7050
   ClientTop       =   480
   ClientWidth     =   4905
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
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
      Left            =   1440
      TabIndex        =   0
      Text            =   " "
      Top             =   3120
      Width           =   2052
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
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4200
      Width           =   2052
   End
   Begin VB.CommandButton cmdSignIn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sign in"
      Height          =   372
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1212
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
      Left            =   480
      TabIndex        =   6
      Top             =   600
      Width           =   3975
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
      Height          =   615
      Left            =   1080
      TabIndex        =   5
      Top             =   5640
      Width           =   2655
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
      Left            =   1440
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
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
      Left            =   1440
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
End
Attribute VB_Name = "mainpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As New Connection
Dim R As New Recordset
Dim S, sql As String
Private Sub cmdClear_Click()
txtUsername.Text = ""
txtPassword.Text = ""
End Sub
Private Sub cmdSignIn_Click()
If Trim(txtUsername.Text) = "admin" And Trim(txtPassword.Text) = "admin" Then
 lebMessage.Visible = False
    Unload Me
    AdminFunctions.Show
Else
C.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\Teacher.mdb;Persist Security Info=False"
sql = "Select * From  TeacherTable where username='" & Trim(txtUsername.Text) & "' and password='" & Trim(txtPassword.Text) & "'"
R.Open sql, C, adOpenDynamic, adLockOptimistic
    If Not R.BOF And Not R.EOF Then
        lebMessage.Visible = False
        Me.Hide
        TeacherFunctions.Show
        TeacherFunctions.txtUsername = Me.txtUsername.Text
    Else
        lebMessage.Caption = "The username or password you entered is incorrect."
        lebMessage.ForeColor = vbRed
        lebMessage.BackColor = vbWhite
        lebMessage.BackStyle = 1
    End If
R.Close
C.Close
End If
End Sub

Private Sub Form_Activate()
txtUsername.SetFocus
End Sub

