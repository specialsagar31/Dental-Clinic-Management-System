VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15570
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   15570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1440
      ScaleHeight     =   270
      ScaleWidth      =   1515
      TabIndex        =   9
      Top             =   6360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5640
      MaxLength       =   5
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3720
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5640
      TabIndex        =   1
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "...Family and Cosmetic Dentistry"
      BeginProperty Font 
         Name            =   "Rage Italic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   8640
      TabIndex        =   7
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "New User? Click here to Sign up"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5040
      MousePointer    =   4  'Icon
      TabIndex        =   5
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Smile Dental Clinic"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1335
      Left            =   3120
      TabIndex        =   6
      Top             =   480
      Width           =   11295
   End
   Begin VB.Image Image2 
      Height          =   5805
      Left            =   8040
      Picture         =   "Form1.frx":0000
      Top             =   2400
      Width           =   7755
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Smile Dental Clinic"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1335
      Left            =   3240
      TabIndex        =   0
      Top             =   480
      Width           =   11295
   End
   Begin VB.Image Image1 
      Height          =   5430
      Left            =   -240
      Picture         =   "Form1.frx":70CB
      Top             =   3000
      Width           =   6300
   End
   Begin VB.Image Image3 
      Height          =   2070
      Left            =   1080
      Picture         =   "Form1.frx":CC24
      Top             =   120
      Width           =   2475
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql1 As String
Private Sub Command1_Click()
Dim m
If Text1.Text = "" Then
MsgBox "Enter Username"
Text1.SetFocus
ElseIf Text2.Text = "" Then
MsgBox "Enter your Password"
Else
sql1 = "select * from Patient_Info"
con.Open "Provider= Microsoft.JET.OLEDB.4.0;Data source= D:\Sagardb.mdb"
rs.Open sql1, con, adOpenDynamic, adLockOptimistic, adCmdText
Do Until rs.EOF
If Text1.Text = rs!UserName And Text2.Text = rs!Password Then
Unload Me
Form3.Show
rs.Close
con.Close
Exit Sub
End If
rs.MoveNext
Loop
m = MsgBox("Login Unsuccessful. You have either entered Username or Password Incorrect. Please Try Again", vbOKCancel, "Message")
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
rs.Close
con.Close
End If
End Sub


Private Sub Form_Load()
Form1.BackColor = RGB(53, 109, 198)
End Sub

Private Sub Label4_Click()
Unload Form1
Form2.Show
End Sub

