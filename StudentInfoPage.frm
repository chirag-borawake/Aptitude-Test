VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form GetStudentDetails 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aptitude Test And Career Guidance"
   ClientHeight    =   7560
   ClientLeft      =   7050
   ClientTop       =   480
   ClientWidth     =   4905
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "FirstName"
      DataSource      =   "StudentADODC"
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
      Left            =   2280
      TabIndex        =   1
      Top             =   2040
      Width           =   1728
   End
   Begin VB.TextBox txtLName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "LastName"
      DataSource      =   "StudentADODC"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   2520
      Width           =   1728
   End
   Begin VB.TextBox txtContactNo 
      BackColor       =   &H00FFFFFF&
      DataField       =   "contactNo"
      DataSource      =   "StudentADODC"
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   4
      Text            =   " "
      Top             =   3480
      Width           =   1728
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H00FFFFFF&
      DataField       =   "username"
      DataSource      =   "StudentADODC"
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
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   " "
      Top             =   4440
      Width           =   1692
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      DataField       =   "password"
      DataSource      =   "StudentADODC"
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
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   " "
      Top             =   4920
      Width           =   1692
   End
   Begin VB.CommandButton cmdRemoveStudent 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete  Student"
      Height          =   492
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1320
      Width           =   1212
   End
   Begin VB.CommandButton cmdViewDetails 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Student Details"
      Height          =   492
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1320
      Width           =   1296
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   0
      Width           =   732
   End
   Begin MSAdodcLib.Adodc StudentADODC 
      Height          =   315
      Left            =   1320
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\Student.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\Student.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "StudentTable"
      Caption         =   "StudentRecords"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   5760
      Width           =   2652
      Begin VB.OptionButton optOther 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Other"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Left            =   1920
         TabIndex        =   22
         Top             =   240
         Width           =   732
      End
      Begin VB.OptionButton opt12th 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "12 th"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   732
      End
      Begin VB.OptionButton opt10th 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "10 th"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   732
      End
   End
   Begin VB.OptionButton OptFemale 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Female"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2880
      TabIndex        =   9
      Top             =   5520
      Width           =   852
   End
   Begin VB.OptionButton optMale 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Male"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   8
      Top             =   5520
      Width           =   972
   End
   Begin VB.TextBox txtEmailID 
      BackColor       =   &H00FFFFFF&
      DataField       =   "mailID"
      DataSource      =   "StudentADODC"
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
      Left            =   2280
      TabIndex        =   5
      Text            =   " "
      Top             =   3960
      Width           =   1692
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox txtStudentAge 
      BackColor       =   &H00FFFFFF&
      DataField       =   "age"
      DataSource      =   "StudentADODC"
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
      Left            =   2280
      TabIndex        =   3
      Text            =   " "
      Top             =   3000
      Width           =   1692
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "First name"
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
      Left            =   1080
      TabIndex        =   28
      Top             =   2040
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last name"
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
      Left            =   1080
      TabIndex        =   27
      Top             =   2520
      Width           =   990
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile number"
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
      Left            =   1080
      TabIndex        =   26
      Top             =   3480
      Width           =   990
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
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
      Left            =   1080
      TabIndex        =   25
      Top             =   4440
      Width           =   990
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
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
      Left            =   1080
      TabIndex        =   24
      Top             =   4920
      Width           =   990
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
      TabIndex        =   23
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Label Label6 
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
      TabIndex        =   21
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Height          =   300
      Left            =   720
      TabIndex        =   20
      Top             =   5880
      Width           =   630
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
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
      Height          =   300
      Left            =   1080
      TabIndex        =   19
      Top             =   3960
      Width           =   990
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Height          =   300
      Left            =   720
      TabIndex        =   18
      Top             =   5520
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
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
      Height          =   300
      Left            =   1080
      TabIndex        =   17
      Top             =   3000
      Width           =   990
   End
End
Attribute VB_Name = "GetStudentDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selStartPW As Integer
Dim selLengthPW As Integer
Dim selStartUN As Integer
Dim selLengthUN As Integer
Private Sub cmdLogOut_Click()
Unload Me
Unload mainpage
mainpage.Show
End Sub

Private Sub cmdNew_Click()
StudentADODC.Recordset.AddNew
txtFName.Text = ""
txtLName.Text = ""
txtStudentAge.Text = ""
txtContactNo.Text = ""
txtEmailID.Text = ""
txtUsername.Text = ""
txtPassword.Text = ""
opt10th.Value = False
opt12th.Value = False
optOther.Value = False
optMale.Value = False
optFemale.Value = False
labMessage.Visible = False
txtFName.SetFocus
End Sub

Private Sub cmdSave_Click()
StudentADODC.Recordset.Fields(1) = txtFName.Text
StudentADODC.Recordset.Fields(2) = txtLName.Text
StudentADODC.Recordset.Fields(3) = txtStudentAge.Text
StudentADODC.Recordset.Fields(4) = txtContactNo.Text
StudentADODC.Recordset.Fields(5) = txtEmailID.Text
If optMale.Value = True Then
    StudentADODC.Recordset.Fields(6) = "Male"
Else
    StudentADODC.Recordset.Fields(6) = "Female"
End If

If opt10th.Value = True Then
    StudentADODC.Recordset.Fields(7) = "10th"
ElseIf opt12th.Value = True Then
    StudentADODC.Recordset.Fields(7) = "12th"
Else
    StudentADODC.Recordset.Fields(7) = "Other"
End If
StudentADODC.Recordset.Fields(8) = txtUsername.Text
StudentADODC.Recordset.Fields(9) = txtPassword.Text
StudentADODC.Recordset.MoveNext
StudentADODC.Recordset.Save
StudentADODC.Refresh
If Not StudentADODC.EOFAction And Not StudentADODC.BOFAction Then
labMessage.Visible = True
labMessage.Caption = " Student Added Successfully!"
labMessage.ForeColor = vbRed
labMessage.BackColor = vbWhite
labMessage.BackStyle = 1
End If
End Sub

Private Sub cmdRemoveStudent_Click()
Unload Me
DeleteStudent.Show
End Sub


Private Sub cmdViewDetails_Click()
Unload Me
StudentDetailsPage.Show
End Sub


Private Sub txtContactNo_Change()
txtPassword.Text = Left$(txtFName.Text, 3) + Left$(txtLName.Text, 3) + Right$(txtContactNo.Text, 3)

 selStartPW = txtPassword.SelStart
    selLengthPW = txtPassword.SelLength
    txtPassword.Text = StrConv(txtPassword.Text, _
    vbLowerCase)
    txtPassword.SelStart = selStartPW
    txtPassword.SelLength = selLengthPW

End Sub

Private Sub txtFName_Change()
txtUsername.Text = txtFName.Text + "." + txtLName.Text

 selStartUN = txtUsername.SelStart
    selLengthUN = txtUsername.SelLength
    txtUsername.Text = StrConv(txtUsername.Text, _
    vbLowerCase)
    txtUsername.SelStart = selStartUN
    txtUsername.SelLength = selLengthUN
    
txtPassword.Text = Left$(txtFName.Text, 3) + Left$(txtLName.Text, 3) + Right$(txtContactNo.Text, 3)

 selStartPW = txtPassword.SelStart
    selLengthPW = txtPassword.SelLength
    txtPassword.Text = StrConv(txtPassword.Text, _
    vbLowerCase)
    txtPassword.SelStart = selStartPW
    txtPassword.SelLength = selLengthPW

End Sub

Private Sub txtLName_Change()
txtUsername.Text = txtFName.Text + "." + txtLName.Text

selStartUN = txtUsername.SelStart
    selLengthUN = txtUsername.SelLength
    txtUsername.Text = StrConv(txtUsername.Text, _
    vbLowerCase)
    txtUsername.SelStart = selStartUN
    txtUsername.SelLength = selLengthUN
txtPassword.Text = Left$(txtFName.Text, 3) + Left$(txtLName.Text, 3) + Right$(txtContactNo.Text, 3)

 selStartPW = txtPassword.SelStart
    selLengthPW = txtPassword.SelLength
    txtPassword.Text = StrConv(txtPassword.Text, _
    vbLowerCase)
    txtPassword.SelStart = selStartPW
    txtPassword.SelLength = selLengthPW

End Sub






Private Sub txtFName_KeyPress(KeyAscii As Integer)
     KeyAscii = alphabetOnly(KeyAscii)
End Sub
Private Sub txtLName_KeyPress(KeyAscii As Integer)
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

Private Sub txtContactNo_KeyPress(KeyAscii As Integer)
    KeyAscii = integerOnly(KeyAscii)
End Sub
Private Sub txtStudentAge_KeyPress(KeyAscii As Integer)
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

Private Sub txtEmailID_KeyPress(KeyAscii As Integer)
KeyAscii = integerAndAlphabets(KeyAscii)
End Sub
Private Function integerAndAlphabets(intKey As Integer) As Integer
  Dim intReturn As Integer
  intReturn = intKey
  Select Case intKey
    Case vbKeyBack        'allow backspace
    Case vbKey0 To vbKey9 'allow 0 to 9
    Case vbKeyA To vbKeyZ 'allow A to Z
    Case 97 To 122        'allow a to z
    Case 64               'allow  @
    Case 46               'allow  .
    Case 95               'allow " _ "
    Case Else             'block all other input
      intReturn = 0
  End Select
  integerAndAlphabets = intReturn
End Function
