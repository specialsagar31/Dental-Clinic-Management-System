VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00404040&
   Caption         =   "Welcome to Smile"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16035
   LinkTopic       =   "Form3"
   ScaleHeight     =   8490
   ScaleWidth      =   16035
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "BOOK NOW"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   8280
      TabIndex        =   6
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "CONSULT NOW"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   5040
      TabIndex        =   5
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Image Image6 
      Height          =   2625
      Left            =   3120
      Picture         =   "Form3.frx":0000
      Top             =   4200
      Width           =   9690
   End
   Begin VB.Image Image4 
      Height          =   720
      Left            =   16440
      Picture         =   "Form3.frx":5690
      Top             =   600
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   7740
      Left            =   -120
      Picture         =   "Form3.frx":8C7A
      Top             =   3480
      Width           =   6210
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   18960
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Enquiry   l"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   16680
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contact   l"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   15360
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Home    l"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   14280
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   16560
      Picture         =   "Form3.frx":1009B
      Top             =   10200
      Width           =   3795
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Smile..."
      BeginProperty Font 
         Name            =   "Brush Script MT"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1455
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   7935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   0
      Top             =   0
      Width           =   22095
   End
   Begin VB.Image Image3 
      Height          =   8295
      Left            =   8880
      Picture         =   "Form3.frx":12633
      Top             =   3480
      Width           =   8580
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   -240
      Top             =   1320
      Width           =   22095
   End
   Begin VB.Image Image5 
      Height          =   5595
      Left            =   -960
      Picture         =   "Form3.frx":1BE27
      Top             =   -2040
      Width           =   23250
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Form3.BackColor = RGB(53, 109, 198)
End Sub

Private Sub Label2_Click()
Form4.Show
End Sub

Private Sub Label4_Click()
Dim MSG As String
MSG = MsgBox("Are You Sure you want to Logout?", vbYesNo, "Logout Confirmation")
If MSG = vbYes Then
   Unload Me
   Form1.Show
End If
End Sub

Private Sub Label5_Click()
Dim MSG As String
MSG = MsgBox("Payment Enquiry needs Administrator Login. Do you want to login as an admin?", vbYesNo, "Requires Admin Login")
If MSG = vbYes Then
Form7.Show
End If
End Sub

Private Sub Label7_Click()
Unload Me
Form5.Show
End Sub

Private Sub Label8_Click()
Unload Me
Form6.Show
End Sub
