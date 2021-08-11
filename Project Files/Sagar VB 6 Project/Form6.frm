VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Booking Form"
   ClientHeight    =   8670
   ClientLeft      =   210
   ClientTop       =   1680
   ClientWidth     =   20400
   LinkTopic       =   "Form6"
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   20400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   12600
      TabIndex        =   48
      Top             =   2760
      Visible         =   0   'False
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
      Left            =   10680
      TabIndex        =   47
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Choose a payment mode"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1935
      Left            =   9840
      TabIndex        =   37
      Top             =   480
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CheckBox Check2 
         Caption         =   "Check1"
         Height          =   135
         Left            =   600
         TabIndex        =   42
         Top             =   1320
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   135
         Left            =   600
         TabIndex        =   41
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Pay using Debit Card"
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
         Left            =   960
         TabIndex        =   40
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash on Consult"
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
         Left            =   960
         TabIndex        =   39
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Method:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton Command4 
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
      Height          =   495
      Left            =   4320
      TabIndex        =   36
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Next"
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
      Left            =   2040
      TabIndex        =   35
      Top             =   8400
      Width           =   1455
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
      Left            =   840
      TabIndex        =   18
      Top             =   6360
      Width           =   6975
   End
   Begin VB.ComboBox Combo5 
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
      Left            =   3480
      TabIndex        =   16
      Text            =   "-SELECT-"
      Top             =   5400
      Width           =   3015
   End
   Begin VB.ComboBox Combo4 
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
      Left            =   840
      TabIndex        =   14
      Text            =   "-SELECT-"
      Top             =   5400
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Text            =   "YYYY"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "Form6.frx":0000
      Left            =   2280
      List            =   "Form6.frx":0002
      TabIndex        =   11
      Text            =   "MM"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Text            =   "DD"
      Top             =   4440
      Width           =   1335
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
      Left            =   5400
      TabIndex        =   8
      Top             =   3600
      Width           =   1095
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
      Left            =   4080
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
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
      Left            =   840
      MaxLength       =   10
      TabIndex        =   5
      Top             =   3600
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
      Left            =   4080
      TabIndex        =   3
      Top             =   2640
      Width           =   2415
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
      Left            =   840
      TabIndex        =   1
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Please give your credentials"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   7095
      Left            =   9840
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox Text8 
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
         Left            =   2280
         TabIndex        =   46
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
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
         Left            =   2880
         TabIndex        =   34
         Top             =   6360
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Proceed"
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
         Left            =   840
         TabIndex        =   33
         Top             =   6360
         Width           =   1335
      End
      Begin VB.TextBox Text7 
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
         Left            =   2280
         TabIndex        =   31
         Top             =   4440
         Width           =   1935
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check1"
         Height          =   135
         Left            =   3360
         TabIndex        =   30
         Top             =   600
         Width           =   195
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check1"
         Height          =   135
         Left            =   1920
         TabIndex        =   29
         Top             =   600
         Width           =   195
      End
      Begin VB.TextBox Text6 
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
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   28
         Top             =   3360
         Width           =   855
      End
      Begin VB.ComboBox Combo7 
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
         Left            =   3720
         TabIndex        =   27
         Text            =   "YYYY"
         Top             =   2640
         Width           =   975
      End
      Begin VB.ComboBox Combo6 
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
         Left            =   2280
         TabIndex        =   26
         Text            =   "MM"
         Top             =   2640
         Width           =   975
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
         Left            =   2280
         MaxLength       =   16
         TabIndex        =   25
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Card Holders Name:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   480
         TabIndex        =   45
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rs. 500/-"
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
         Height          =   255
         Left            =   2400
         TabIndex        =   44
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Consultation Fees:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   43
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Image Image3 
         Height          =   690
         Left            =   960
         Picture         =   "Form6.frx":0004
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   3360
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "*Enter Captcha above"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Image Image2 
         Height          =   1035
         Left            =   3600
         Picture         =   "Form6.frx":1A21
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1560
      End
      Begin VB.Image Image1 
         Height          =   1065
         Left            =   2160
         Picture         =   "Form6.frx":693D
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Type the text shown below:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   24
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "ATM PIN:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Card Number:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Card Type:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   495
      Left            =   17640
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label24 
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
      Left            =   17760
      TabIndex        =   52
      Top             =   720
      Width           =   1935
   End
   Begin VB.Image Image7 
      Height          =   9780
      Left            =   6600
      Picture         =   "Form6.frx":8010
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   14235
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Details..."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   600
      TabIndex        =   51
      Top             =   600
      Width           =   6495
   End
   Begin VB.Label Label22 
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
      TabIndex        =   50
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Details..."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   720
      TabIndex        =   49
      Top             =   600
      Width           =   6495
   End
   Begin VB.Image Image6 
      Height          =   5820
      Left            =   15120
      Picture         =   "Form6.frx":7E1AA
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.Image Image4 
      Height          =   5820
      Left            =   15120
      Picture         =   "Form6.frx":88C53
      Stretch         =   -1  'True
      Top             =   5280
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Visible         =   0   'False
      X1              =   9000
      X2              =   9000
      Y1              =   240
      Y2              =   10080
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
      Left            =   840
      TabIndex        =   17
      Top             =   6000
      Width           =   3735
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
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   5040
      Width           =   2295
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
      Left            =   840
      TabIndex        =   13
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Date:"
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
      Left            =   840
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
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
      Left            =   4080
      TabIndex        =   6
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile number:"
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
      Left            =   840
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
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
      Left            =   4080
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
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
      Left            =   840
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Image Image5 
      Height          =   5475
      Left            =   -1680
      Picture         =   "Form6.frx":944D3
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   23250
   End
   Begin VB.Image Image8 
      Height          =   1035
      Left            =   -960
      Picture         =   "Form6.frx":A22D6
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   23250
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql1 As String
Private Sub Check1_Click()
If Check1.Value = 1 Then
Check2.Value = False
Command5.Visible = True
Command6.Visible = True
Image6.Visible = True
ElseIf Check1.Value = 0 Then
Command5.Visible = False
Command6.Visible = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Check1.Value = False
Frame1.Visible = True
Image4.Visible = True
Image6.Visible = False
Else
Frame1.Visible = False
Image4.Visible = False
Image6.Visible = True
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Check4.Value = False
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
Check3.Value = False
End If
End Sub

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
Private Sub Combo3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 8
Case Else
KeyAscii = 0
End Select
End Sub
Private Sub Combo4_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 8
Case Else
KeyAscii = 0
End Select
End Sub
Private Sub Combo5_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 8
Case Else
KeyAscii = 0
End Select
End Sub
Private Sub Combo6_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 8
Case Else
KeyAscii = 0
End Select
End Sub
Private Sub Combo7_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Command1_Click()
If Check3.Value = False And Check4.Value = False Then
MsgBox "Please choose your card type."
ElseIf Text8.Text = "" Then
MsgBox "Enter the name as on your debit card."
ElseIf Text4.Text = "" Then
MsgBox "Enter your card number."
ElseIf Len(Trim(Text4.Text)) <> 16 Then
MsgBox "Please enter your 16-digit card number"
ElseIf Combo6.Text = "MM" Then
MsgBox "Please choose expiry month"
ElseIf Combo7.Text = "YYYY" Then
MsgBox "Please choose expiry year"
ElseIf Text6.Text = "" Then
MsgBox "Please enter your 4-digit ATM PIN"
ElseIf Len(Trim(Text6.Text)) < 4 Then
MsgBox "Please enter correct 4-digit ATM PIN"
ElseIf Text7.Text <> "smwm" Then
MsgBox "Invalid Captcha"
Else
sql1 = "select*from Payment_Info"
con.Open "Provider=Microsoft.JET.OLEDB.4.0;data source=d:\Sagardb.mdb"
rs.Open sql1, con, adOpenDynamic, adLockOptimistic, adCmdText
rs.AddNew
rs!Name_of_Payee = Text8.Text
rs!Payment = "PAID"
rs.Update
rs.Close
con.Close
Dim MSG As String
MSG = MsgBox("Thanks for using your debit card. Click on YES to confirm your booking.", vbYesNo, "Payment Successful")
If MSG = vbYes Then

sql1 = "select*from Booking_Info"
con.Open "Provider=Microsoft.JET.OLEDB.4.0;data source=d:\Sagardb.mdb"
rs.Open sql1, con, adOpenDynamic, adLockOptimistic, adCmdText
rs.AddNew
rs!Fullname = Text1.Text
rs!Email = Text2.Text
rs!Mobile = Text3.Text
If Option1.Value = True Then
rs!Gender = "Male"
ElseIf Option2.Value = True Then
rs!Gender = "Female"
End If
rs!Day1 = Combo1.Text
rs!Month1 = Combo2.Text
rs!Year1 = Combo3.Text
rs!Consult_Time = Combo4.Text
rs!Problem = Combo5.Text
rs!Reason = Text5.Text
rs.Update
MsgBox ("Dear" & " " & Text1.Text & " your Booking has been confirmed. Be ready for your consultation on " & Combo1.Text & "-" & Combo2.Text & "-" & Combo3.Text & " " & " on " & Combo4.Text)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Option1.Value = False
Option2.Value = False
rs.Close
con.Close
Form3.Show
Unload Me
End If
End If
End Sub

Private Sub Command2_Click()
Dim MSG As String
MSG = MsgBox("Are you sure you want to cancel your booking?", vbYesNo, "Cancel")
If MSG = vbYes Then
Form3.Show
Unload Me
End If
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Then
MsgBox "Please enter Full Name.", vbOKOnly + vbExclamation, "Message"
Text1.SetFocus
ElseIf Text2.Text = "" Then
MsgBox "Please enter your email address.", vbOKOnly + vbExclamation, "Message"
Text2.SetFocus
ElseIf Text3.Text = "" Then
MsgBox "Please enter your mobile.", vbOKOnly + vbExclamation, "Message"
Text3.SetFocus
ElseIf Len(Trim(Text3.Text)) < 10 Then
MsgBox "Please enter correct 10-digit Mobile Number", vbOKOnly + vbExclamation, "Message"
Text3.SetFocus
ElseIf Option1.Value = False And Option2.Value = False Then
MsgBox "Please choose your Gender.", vbOKOnly + vbExclamation, "Message"
ElseIf Combo1.Text = "DD" Then
MsgBox "Please choose a day.", vbOKOnly + vbExclamation, "Message"
ElseIf Combo2.Text = "MM" Then
MsgBox "Please choose a month.", vbOKOnly + vbExclamation, "Message"
ElseIf Combo3.Text = "YYYY" Then
MsgBox "Please choose year.", vbOKOnly + vbExclamation, "Message"
ElseIf Combo4.Text = "-SELECT-" Then
MsgBox "Please consultation time.", vbOKOnly + vbExclamation, "Message"
ElseIf Combo5.Text = "-SELECT-" Then
MsgBox "Please choose your problem.", vbOKOnly + vbExclamation, "Message"
Else: Frame2.Visible = True
Line1.Visible = True
Image6.Visible = True
Image7.Visible = False
End If
End Sub

Private Sub Command4_Click()
Dim MSG As String
MSG = MsgBox("Are you sure you want to cancel your booking?", vbYesNo, "Cancel")
If MSG = vbYes Then
Form3.Show
Unload Me
End If
End Sub

Private Sub Command5_Click()
sql1 = "select*from Booking_Info"
con.Open "Provider=Microsoft.JET.OLEDB.4.0;data source=d:\Sagardb.mdb"
rs.Open sql1, con, adOpenDynamic, adLockOptimistic, adCmdText
rs.AddNew
rs!Fullname = Text1.Text
rs!Email = Text2.Text
rs!Mobile = Text3.Text
If Option1.Value = True Then
rs!Gender = "Male"
ElseIf Option2.Value = True Then
rs!Gender = "Female"
End If
rs!Day1 = Combo1.Text
rs!Month1 = Combo2.Text
rs!Year1 = Combo3.Text
rs!Consult_Time = Combo4.Text
rs!Problem = Combo5.Text
rs!Reason = Text5.Text
rs.Update
MsgBox ("Dear" & " " & Text1.Text & " your Booking has been confirmed. Be ready for your consultation with additional charges on " & Combo1.Text & "-" & Combo2.Text & "-" & Combo3.Text & " " & " on " & Combo4.Text)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Option1.Value = False
Option2.Value = False
rs.Close
con.Close
Form3.Show
Unload Me
End Sub

Private Sub Command6_Click()
Dim MSG As String
MSG = MsgBox("Are you sure you want to cancel your booking?", vbYesNo, "Cancel")
If MSG = vbYes Then
Form3.Show
Unload Me
End If
End Sub

Private Sub Form_Load()
Form6.BackColor = RGB(53, 109, 198)
Frame1.BackColor = RGB(53, 109, 198)
Frame2.BackColor = RGB(53, 109, 198)
Combo1.AddItem "DD"
Combo1.AddItem 1
Combo1.AddItem 2
Combo1.AddItem 3
Combo1.AddItem 4
Combo1.AddItem 5
Combo1.AddItem 6
Combo1.AddItem 7
Combo1.AddItem 8
Combo1.AddItem 9
Combo1.AddItem 10
Combo1.AddItem 11
Combo1.AddItem 12
Combo1.AddItem 13
Combo1.AddItem 14
Combo1.AddItem 15
Combo1.AddItem 16
Combo1.AddItem 17
Combo1.AddItem 18
Combo1.AddItem 19
Combo1.AddItem 20
Combo1.AddItem 21
Combo1.AddItem 22
Combo1.AddItem 23
Combo1.AddItem 24
Combo1.AddItem 25
Combo1.AddItem 26
Combo1.AddItem 27
Combo1.AddItem 28
Combo1.AddItem 29
Combo1.AddItem 30
Combo1.AddItem 31
Combo2.AddItem "January"
Combo2.AddItem "February"
Combo2.AddItem "March"
Combo2.AddItem "April"
Combo2.AddItem "May"
Combo2.AddItem "June"
Combo2.AddItem "July"
Combo2.AddItem "August"
Combo2.AddItem "September"
Combo2.AddItem "October"
Combo2.AddItem "November"
Combo2.AddItem "December"
Combo3.AddItem 2019
Combo3.AddItem 2020
Combo4.AddItem "10:30am-11:00am"
Combo4.AddItem "11:00am-11:30am"
Combo4.AddItem "11:30am-12:00pm"
Combo4.AddItem "12:00pm-12:30pm"
Combo4.AddItem "12:30pm-01:00pm"
Combo4.AddItem "01:30pm-01:30pm"
Combo5.AddItem "-SELECT-"
Combo5.AddItem "Inlays & Onlays"
Combo5.AddItem "Root Canal Treatment"
Combo5.AddItem "Teeth Whitening"
Combo5.AddItem "Complete Dentures"
Combo5.AddItem "Cosmetic Dentistry"
Combo5.AddItem "Crowns & Bridges"
Combo5.AddItem "Dental Implants"
Combo5.AddItem "Oral & Maxillofacial Surgery"
Combo5.AddItem "Orthodontic Treatment"
Combo5.AddItem "Other general treatments"
Combo6.AddItem "MM"
Combo6.AddItem 1
Combo6.AddItem 2
Combo6.AddItem 3
Combo6.AddItem 4
Combo6.AddItem 5
Combo6.AddItem 6
Combo6.AddItem 7
Combo6.AddItem 8
Combo6.AddItem 9
Combo6.AddItem 10
Combo6.AddItem 11
Combo6.AddItem 12
Combo7.AddItem 2019
Combo7.AddItem 2020
Combo7.AddItem 2021
Combo7.AddItem 2022
Combo7.AddItem 2023
Combo7.AddItem 2024
Combo7.AddItem 2025
Combo7.AddItem 2026
Combo7.AddItem 2027
Combo7.AddItem 2028
Combo7.AddItem 2029
Combo7.AddItem 2030
Combo7.AddItem 2031
Combo7.AddItem 2032
Combo7.AddItem 2033
Combo7.AddItem 2034
Combo7.AddItem 2035
End Sub

Private Sub Label24_Click()
Dim MSG As String
MSG = MsgBox("Are you sure you want to cancel your booking?", vbYesNo, "Cancel")
If MSG = vbYes Then
Form3.Show
Unload Me
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 65 To 90 'uppercase
Case 97 To 122 'lowercase
Case 32 'spacebar
Case 8 'backspace
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case 8
Case Else
KeyAscii = 0
End Select
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 65 To 90 'uppercase
Case 97 To 122 'lowercase
Case 32 'spacebar
Case 8 'backspace
Case Else
KeyAscii = 0
End Select
End Sub
