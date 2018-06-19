VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form ResultPage 
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
      Height          =   3495
      Left            =   360
      OleObjectBlob   =   "ResultPage.frx":0000
      TabIndex        =   7
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label labTimeUsed 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   5880
      Width           =   555
   End
   Begin VB.Label labWrongAnswer 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label labCorrectAnswer 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Correct answer"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   4440
      Width           =   1830
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Wrong answer"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   5160
      Width           =   1830
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Time used"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   5880
      Width           =   1350
   End
End
Attribute VB_Name = "ResultPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim result(1 To 3, 1 To 3) As Variant

Private Sub cmdNext_Click()
Unload Me
chart.Show
End Sub

Private Sub Form_Activate()
result(1, 2) = "Correct"
result(1, 3) = "Wroung"
result(2, 1) = " "
result(2, 2) = Val(labCorrectAnswer.Caption)
result(2, 3) = Val(labWrongAnswer.Caption)
MSChart1.ChartData = result
MSChart1.ShowLegend = True
MSChart1.chartType = VtChChartType2dPie
MSChart1.RowCount = 1
MSChart1.ColumnCount = 2
MSChart1.Plot.SeriesCollection(1).DataPoints(-1).Brush.FillColor.Set 2, 250, 2
MSChart1.Plot.SeriesCollection(2).DataPoints(-1).Brush.FillColor.Set 250, 2, 2

MSChart1.Legend.VtFont.VtColor.Set 210, 210, 210
MSChart1.Legend.VtFont.Size = 10
End Sub



