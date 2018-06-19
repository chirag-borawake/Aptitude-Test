VERSION 5.00
Begin VB.Form DeleteStudent 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aptitude Test And Career Guidance"
   ClientHeight    =   7560
   ClientLeft      =   7050
   ClientTop       =   480
   ClientWidth     =   4905
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   732
   End
   Begin VB.CommandButton cmdAddStudent 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Student"
      Height          =   492
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1212
   End
   Begin VB.CommandButton cmdViewDetails 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Student Details"
      Height          =   492
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1296
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      Height          =   372
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   1212
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
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   2760
      TabIndex        =   0
      Top             =   3720
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   5160
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
      Left            =   720
      TabIndex        =   6
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter student id which is to be deleted"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   3720
      Width           =   1815
   End
End
Attribute VB_Name = "DeleteStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim C As New Connection
Dim R As New Recordset
Dim S As String

Private Sub cmdAddStudent_Click()
Unload Me
GetStudentDetails.Show
End Sub

Private Sub cmdCancel_Click()
Unload Me
TeacherFunctions.Show
End Sub


Private Sub cmdDelete_Click()
Dim id As Integer
id = Val(txtDelete.Text)
C.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\Student.mdb;Persist Security Info=False"
S = "Select * from StudentTable Where(Id = " & id & ")"
R.Open S, C, adOpenDynamic, adLockBatchOptimistic
If Not R.EOF And Not R.BOF Then
    R.Close
    S = "Delete from StudentTable where (Id=" & id & ")"
    R.Open S, C, adOpenDynamic, adLockOptimistic
    labMessage.Caption = "Record is deleted."
    labMessage.ForeColor = vbRed
    labMessage.BackColor = vbWhite
    labMessage.BackStyle = 1
Else
    
    labMessage.Caption = "Record is not present."
    labMessage.ForeColor = vbRed
    labMessage.BackColor = vbWhite
    labMessage.BackStyle = 1
End If
C.Close
End Sub

Private Sub cmdLogOut_Click()
Unload Me
Unload mainpage
mainpage.Show
End Sub

Private Sub cmdViewDetails_Click()
Unload Me
StudentDetailsPage.Show
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

