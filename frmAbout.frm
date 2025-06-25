VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About CoLIN"
   ClientHeight    =   3060
   ClientLeft      =   3690
   ClientTop       =   4605
   ClientWidth     =   5400
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReadme 
      Caption         =   "Readme"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "http://www.barc0de.demon.co.uk/nlp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "alan@barc0de.demon.co.uk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Copyright 2000 alan j. brown"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label version 
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "CoLIN - Computer Linguistic ImitatioN"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":0316
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5280
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   5280
      X2              =   120
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib _
    "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long



Private Const SW_SHOWNORMAL = 1

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdReadme_Click()
cinWeb App.Path & "\readme.txt"
End Sub

Private Sub Form_Load()
version = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
Public Sub cinWeb(htmlfile, Optional dr)
If IsMissing(dr) Then
    dr = "C:\"
End If
        Dim iret As Long
            iret = ShellExecute(frmAbout.hwnd, _
            vbNullString, _
            htmlfile, _
            vbNullString, _
            dr, _
            SW_SHOWNORMAL)
End Sub

Private Sub Label4_Click()
cinWeb "mailto:alan@barcode.demon.co.uk"
End Sub

Private Sub Label5_Click()
cinWeb "http://www.barc0de.demon.co.uk/nlp"
End Sub
