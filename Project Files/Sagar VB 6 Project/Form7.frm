VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00404040&
   Caption         =   "Administrator Login"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10815
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Cancel"
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
      Left            =   5160
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Submit"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   3600
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3600
      TabIndex        =   2
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Portal"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   2520
      TabIndex        =   6
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label1 
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
      Left            =   2160
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Admin:"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   3105
      Left            =   7080
      Picture         =   "Form7.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   3060
   End
   Begin VB.Image Image5 
      Height          =   9795
      Left            =   -1320
      Picture         =   "Form7.frx":0F2B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   23250
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command5_Click()
Dim MSG As String
Dim SG As String
If Text1.Text <> "Sagar@12345" Then
MSG = MsgBox("Invalid Admin!", vbRetryCancel + vbCritical, "Message")
  If MSG = vbRetry Then
  Text1.SetFocus
  Text1.Text = ""
  Text2.Text = ""
  End If
ElseIf Text2.Text <> "12345" Then
SG = MsgBox("Password Incorrect", vbCritical + vbRetryCancel, "Message")
    If SG = vbRetry Then
    Text1.SetFocus
    Text1.Text = ""
    Text2.Text = ""
    End If
ElseIf MsgBox("You have Logged in Successfully", vbOKOnly, "Login Successful") Then
Form8.Show
Form3.Hide
Unload Me
End If
End Sub

Private Sub Command6_Click()
Dim MSG As String
MSG = MsgBox("Are you sure you want to cancel?", vbYesNo, "Cancel")
If MSG = vbYes Then
Unload Me
Form3.Show
End If
End Sub

Private Sub Form_Load()
Form7.BackColor = RGB(53, 109, 198)
End Sub

