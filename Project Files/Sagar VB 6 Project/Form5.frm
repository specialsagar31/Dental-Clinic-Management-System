VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00404040&
   Caption         =   "Consultation Form"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15720
   LinkTopic       =   "Form5"
   ScaleHeight     =   8835
   ScaleWidth      =   15720
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   1635
      TabIndex        =   24
      Top             =   9960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   21
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
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
      Height          =   495
      Left            =   2280
      TabIndex        =   20
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   7800
      Width           =   195
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   15
      Text            =   "-SELECT-"
      Top             =   4200
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Text            =   "-SELECT-"
      Top             =   4200
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C000&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   720
      TabIndex        =   11
      Top             =   5400
      Width           =   6975
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      MaxLength       =   10
      TabIndex        =   10
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   1320
      Width           =   2415
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H008080FF&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A Smile can change your day!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   10800
      TabIndex        =   23
      Top             =   1680
      Width           =   8775
   End
   Begin VB.Image Image2 
      Height          =   7230
      Left            =   11760
      Picture         =   "Form5.frx":0000
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   8430
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   495
      Left            =   17880
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Back to Homepage"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   18000
      TabIndex        =   22
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "I declare that all the information above is true as per my knowledge and I have no problem in sharing this information."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   1440
      TabIndex        =   18
      Top             =   7800
      Width           =   6855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "*Please fill up these credentials:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose your problem:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5640
      TabIndex        =   16
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Consult time:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   13
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Reason for Consultation(optional):"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   5040
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Number:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   16275
      Left            =   -1920
      Picture         =   "Form5.frx":3E896
      Stretch         =   -1  'True
      Top             =   0
      Width           =   23250
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql1 As String

Private Sub Combo1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Please enter Full Name", vbOKOnly + vbExclamation, "Message"
Text1.SetFocus
ElseIf Text2.Text = "" Then
MsgBox "Please enter your email address", vbOKOnly + vbExclamation, "Message"
Text2.SetFocus
ElseIf Text3.Text = "" Then
MsgBox "Please choose your country", vbOKOnly + vbExclamation, "Message"
Text3.SetFocus
ElseIf Text4.Text = "" Then
MsgBox "Please provide your contact number", vbOKOnly + vbExclamation, "Message"
Text4.SetFocus
ElseIf Len(Trim(Text4.Text)) <> 10 Then
MsgBox "Please enter correct 10-digit mobile number", vbOKOnly + vbExclamation, "Message"
ElseIf Option1.Value = False And Option2.Value = False Then
MsgBox "Please choose your Gender", vbOKOnly + vbExclamation, "Message"
ElseIf Combo1.Text = "-SELECT-" Then
MsgBox "Please choose a time period to consult", vbOKOnly + vbExclamation, "Message"
ElseIf Combo2.Text = "-SELECT-" Then
MsgBox "Please choose a problem to consult", vbOKOnly + vbExclamation, "Message"
ElseIf Check1 = False Then
MsgBox "You have missed something", vbOKOnly + vbExclamation, "Message"
Else
sql1 = "select*from Consult_Info"
con.Open "Provider=Microsoft.JET.OLEDB.4.0;data source=d:\Sagardb.mdb"
rs.Open sql1, con, adOpenDynamic, adLockOptimistic, adCmdText
rs.AddNew
rs!Fullname = Text1.Text
rs!Email = Text2.Text
rs!Country = Text3.Text
rs!Mobile = Text4.Text
If Option1.Value = True Then
rs!Gender = "Male"
ElseIf Option2.Value = True Then
rs!Gender = "Female"
End If
rs!Consult_Time = Combo1.Text
rs!Problem = Combo2.Text
rs!Reason = Text5.Text
rs.Update
MsgBox "Dear" & " " & Text1.Text & " your consultation has been set on " & Combo1.Text, vbInformation, "Successfully registered"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Option1.Value = False
Option2.Value = False
rs.Close
con.Close
Form3.Show
Unload Me
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Option1 = False
Option2 = False
Combo1.Text = "-SELECT-"
Combo2.Text = "-SELECT-"
Text5.Text = ""
Check1 = False
Text1.SetFocus
End Sub

Private Sub Combo2_DropDown()
Combo2.AddItem "-SELECT-"
Combo2.AddItem "Inlays & Onlays"
Combo2.AddItem "Root Canal Treatment"
Combo2.AddItem "Teeth Whitening"
Combo2.AddItem "Complete Dentures"
Combo2.AddItem "Cosmetic Dentistry"
Combo2.AddItem "Crowns & Bridges"
Combo2.AddItem "Dental Implants"
Combo2.AddItem "Oral & Maxillofacial Surgery"
Combo2.AddItem "Orthodontic Treatment"
Combo2.AddItem "Other general treatments"
End Sub

Private Sub Form_Load()
Form5.BackColor = RGB(53, 109, 198)
Combo1.AddItem "-SELECT-"
Combo1.AddItem "10:30am -11:00am"
Combo1.AddItem "11:00am -11:30am"
Combo1.AddItem "11:30am -12:00pm"
Combo1.AddItem "12:00pm -12:30pm"
Combo1.AddItem "12:30pm -01:00pm"
End Sub


Private Sub Label11_Click()
Dim MSG As String
MSG = MsgBox("Are you sure you want to cancel your consultation?", vbYesNo, "Cancel")
If MSG = vbYes Then
Form3.Show
Unload Me
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case 8
Case Else
KeyAscii = 0
End Select
End Sub
