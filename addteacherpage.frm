VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form AddTeacherPage 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aptitude Test And Career Guidance"
   ClientHeight    =   7560
   ClientLeft      =   7050
   ClientTop       =   480
   ClientWidth     =   4905
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMailID 
      BackColor       =   &H00FFFFFF&
      DataField       =   "email"
      DataSource      =   "TeacherDataAdodc"
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
      MaxLength       =   10
      TabIndex        =   5
      Text            =   " "
      Top             =   4200
      Width           =   2205
   End
   Begin VB.CommandButton cmdViewQuestions 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Questions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1440
      Width           =   930
   End
   Begin VB.CommandButton cmdViewTeacherDetails 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Teacher Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1440
      Width           =   1170
   End
   Begin VB.CommandButton cmdRemoveTeacher 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete Teacher"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1440
      Width           =   804
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
      Height          =   495
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1440
      Width           =   804
   End
   Begin VB.CommandButton cmdDeleteQuestion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete Question"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1440
      Width           =   810
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
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   0
      Width           =   732
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
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   " "
      Top             =   6120
      Width           =   2175
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
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   " "
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox txtAge 
      BackColor       =   &H00FFFFFF&
      DataField       =   "age"
      DataSource      =   "TeacherDataAdodc"
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
      MaxLength       =   3
      TabIndex        =   6
      Top             =   4680
      Width           =   2205
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      Height          =   372
      Left            =   2880
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6960
      Width           =   972
   End
   Begin MSAdodcLib.Adodc TeacherDataAdodc 
      Height          =   270
      Left            =   1200
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   476
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
      BackColor       =   14737632
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\Teacher.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\Teacher.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TeacherTable"
      Caption         =   "Teacher Records"
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
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New"
      Height          =   372
      Left            =   960
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   972
   End
   Begin VB.TextBox txtContactNo 
      BackColor       =   &H00FFFFFF&
      DataField       =   "contactNo"
      DataSource      =   "TeacherDataAdodc"
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
      MaxLength       =   10
      TabIndex        =   4
      Text            =   " "
      Top             =   3720
      Width           =   2205
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H00FFFFFF&
      DataField       =   "address"
      DataSource      =   "TeacherDataAdodc"
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
      TabIndex        =   3
      Text            =   " "
      Top             =   3240
      Width           =   2205
   End
   Begin VB.TextBox txtLName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "LastName"
      DataSource      =   "TeacherDataAdodc"
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
      TabIndex        =   2
      Text            =   " "
      Top             =   2760
      Width           =   2205
   End
   Begin VB.OptionButton optFemale 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Female"
      ForeColor       =   &H8000000F&
      Height          =   192
      Left            =   2880
      TabIndex        =   8
      Top             =   5160
      Width           =   972
   End
   Begin VB.OptionButton optMale 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Male"
      ForeColor       =   &H8000000F&
      Height          =   192
      Left            =   2040
      TabIndex        =   7
      Top             =   5160
      Width           =   732
   End
   Begin VB.TextBox txtFName 
      BackColor       =   &H00FFFFFF&
      DataField       =   "FirstName"
      DataSource      =   "TeacherDataAdodc"
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
      TabIndex        =   1
      Text            =   " "
      Top             =   2280
      Width           =   2205
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
      TabIndex        =   28
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Aptitude Test And Career Guidance"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Left            =   960
      TabIndex        =   27
      Top             =   360
      Width           =   3015
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
      Height          =   372
      Left            =   720
      TabIndex        =   26
      Top             =   6120
      Width           =   996
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
      Height          =   372
      Left            =   720
      TabIndex        =   25
      Top             =   5640
      Width           =   996
   End
   Begin VB.Label Label7 
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
      ForeColor       =   &H8000000F&
      Height          =   252
      Left            =   720
      TabIndex        =   24
      Top             =   5160
      Width           =   996
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H8000000F&
      Height          =   252
      Left            =   720
      TabIndex        =   23
      Top             =   4680
      Width           =   996
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H8000000F&
      Height          =   252
      Left            =   720
      TabIndex        =   22
      Top             =   4200
      Width           =   996
   End
   Begin VB.Label Label4 
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
      Height          =   252
      Left            =   720
      TabIndex        =   21
      Top             =   3720
      Width           =   996
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Height          =   252
      Left            =   720
      TabIndex        =   20
      Top             =   3240
      Width           =   996
   End
   Begin VB.Label Label2 
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
      Height          =   252
      Left            =   720
      TabIndex        =   19
      Top             =   2760
      Width           =   996
   End
   Begin VB.Label Label1 
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
      Height          =   252
      Left            =   720
      TabIndex        =   18
      Top             =   2280
      Width           =   996
   End
End
Attribute VB_Name = "AddTeacherPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selStartPW As Integer
Dim selLengthPW As Integer
Dim selStartUN As Integer
Dim selLengthUN As Integer
Dim gender As String
Private Sub CancelButton_Click()
Unload Me
AdminFunctions.Show
End Sub

Private Sub cmdAdd_Click()
TeacherDataAdodc.Recordset.AddNew
txtFName.Text = ""
txtLName.Text = ""
txtAddress.Text = ""
txtContactNo.Text = ""
txtMailID.Text = ""
txtAge.Text = ""
txtUsername.Text = ""
txtPassword.Text = ""
optFemale.Value = False
optMale.Value = False
optFemale.Enabled = True
optMale.Enabled = True
labMessage.Visible = False
txtFName.SetFocus
End Sub

Private Sub cmdAddQuestion_Click()
Unload Me
QuestionScreen.Show
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

Private Sub cmdSave_Click()
TeacherDataAdodc.Recordset.Fields(1) = txtFName.Text
TeacherDataAdodc.Recordset.Fields(2) = txtLName.Text
TeacherDataAdodc.Recordset.Fields(3) = txtAddress.Text
TeacherDataAdodc.Recordset.Fields(4) = txtContactNo.Text
TeacherDataAdodc.Recordset.Fields(5) = txtMailID.Text
TeacherDataAdodc.Recordset.Fields(6) = txtAge.Text
TeacherDataAdodc.Recordset.Fields(7) = gender
TeacherDataAdodc.Recordset.Fields(8) = txtUsername.Text
TeacherDataAdodc.Recordset.Fields(9) = txtPassword.Text
TeacherDataAdodc.Recordset.MoveNext
TeacherDataAdodc.Refresh
If Not TeacherDataAdodc.Recordset.EOF And Not TeacherDataAdodc.Recordset.BOF Then
labMessage.Visible = True
labMessage.Caption = " Teacher Added Successfully!"
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


Private Sub optFemale_Click()
gender = "Female"
End Sub

Private Sub optMale_Click()
gender = "Male"
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
Private Sub txtAge_KeyPress(KeyAscii As Integer)
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

Private Sub txtMailID_KeyPress(KeyAscii As Integer)
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
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
     KeyAscii = integerAndAlphabetsSymbols(KeyAscii)
End Sub
Private Function integerAndAlphabetsSymbols(intKey As Integer) As Integer
  Dim intReturn As Integer
  intReturn = intKey
  Select Case intKey
    Case vbKeyBack        'allow backspace
    Case vbKey0 To vbKey9 'allow 0 to 9
    Case vbKeyA To vbKeyZ 'allow A to Z
    Case 97 To 122        'allow a to z
    Case 58                'allow  :
    Case 59                'allow  ;
    Case 44 To 46          'allow  . , -
    Case 95               'allow " _"
    Case 32               'allow "  "
    Case 34               'allow "
    Case Else             'block all other input
      intReturn = 0
  End Select
  integerAndAlphabetsSymbols = intReturn
End Function

