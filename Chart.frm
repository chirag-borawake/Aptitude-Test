VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form chart 
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
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5295
      Left            =   360
      OleObjectBlob   =   "Chart.frx":0000
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox txtMedical 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   7200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtArts 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   7200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtFinance 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtTechnical 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   7200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exit"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Career Graph"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   2670
   End
End
Attribute VB_Name = "chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
Unload Me
Unload mainpage
mainpage.Show
End Sub

Private Sub Form_Activate()
MSChart1.chartType = VtChChartType2dBar
MSChart1.RowCount = 4
Dim fieldCount As Integer
fieldCount = 1
 For i = 1 To 4
    MSChart1.Row = i
    MSChart1.Column = 1
    If i = 1 Then
        MSChart1.RowLabel = "Technical"
        MSChart1.Data = Val(txtTechnical.Text)
    ElseIf i = 2 Then
        MSChart1.RowLabel = "Medical"
         MSChart1.Data = Val(txtMedical.Text)
    ElseIf i = 3 Then
        MSChart1.RowLabel = "Arts"
        MSChart1.Data = Val(txtArts.Text)
    Else
        MSChart1.RowLabel = "Finance"
        MSChart1.Data = Val(txtFinance.Text)
    End If
Next
End Sub
 

