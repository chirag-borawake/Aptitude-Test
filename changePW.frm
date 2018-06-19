VERSION 5.00
Begin VB.Form changePW 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chandge pass word"
   ClientHeight    =   7410
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   4755
   Begin VB.TextBox txtCurrentPW 
      Height          =   372
      Left            =   2280
      TabIndex        =   5
      Top             =   960
      Width           =   972
   End
   Begin VB.TextBox txtConfermPW 
      Height          =   372
      Left            =   2280
      TabIndex        =   4
      Top             =   2280
      Width           =   972
   End
   Begin VB.TextBox txtNewPW 
      Height          =   372
      Left            =   2280
      TabIndex        =   2
      Top             =   1560
      Width           =   972
   End
   Begin VB.CommandButton cmdChangePW 
      Caption         =   "Change"
      Height          =   372
      Left            =   2280
      TabIndex        =   0
      Top             =   3000
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   720
      TabIndex        =   3
      Top             =   2160
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "New password"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "changePW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As New ADODB.Connection
Dim R As New ADODB.Recordset
Dim S As String
Dim sql As String
Private Sub cmdChangePW_Click()
C.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Chirag\temp\project\Teacher.mdb;Persist Security Info=False"
'S = "Update ProductTable Set pcode = '" & txtPname.Text & "', pname = '" & txtPname.Text & "',quantity=" & Val(txtQty.Text) & ",price=" & Val(txtPrice.Text) & "',tprice=" & Val(txtTotal.Text) & " Where pcode = " & Val(txtPcode.Text)"
sql = "update TeacherTable set password='" & txtNewPW.Text & "'," _
& "username='" & txtCurrentPW.Text & "'" _
& "where ID ='" & txtConfermPW.Text & "'"
Set R = C.Execute(sql)
Set R = Nothing
'R.Open sql, C, adOpenDynamic, adLockOptimistic
'C.Close
'Set R = C.Execute(sql)
'Set R = Nothing
MsgBox "Updated"
End Sub

