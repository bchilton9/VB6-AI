VERSION 5.00
Begin VB.Form nlm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CoLIN - Computer Linguistic ImitatioN"
   ClientHeight    =   5415
   ClientLeft      =   4200
   ClientTop       =   4200
   ClientWidth     =   9345
   Icon            =   "nlp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5415
   ScaleWidth      =   9345
   Begin VB.PictureBox picprog 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   11
      Top             =   3840
      Width           =   5895
      Begin VB.Shape pbsp 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Rate Reply"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   4935
      Begin VB.OptionButton optRate 
         Caption         =   "Shit"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optRate 
         Caption         =   "Bad"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optRate 
         Caption         =   "Ok"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optRate 
         Caption         =   "Good"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optRate 
         Caption         =   "Fantastic"
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "CoLIN"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4935
      Begin VB.TextBox Computer 
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Human"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      Begin VB.TextBox Human 
         Height          =   1365
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Data datWords 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "App.Path & ""\nlp.mdb"""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Words"
      Top             =   240
      Width           =   2700
   End
   Begin VB.CommandButton Talk 
      Caption         =   "Talk"
      Default         =   -1  'True
      Height          =   3615
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear Data"
      End
      Begin VB.Menu mnucap1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuChuman 
         Caption         =   "Copy &Human"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCcomp 
         Caption         =   "Copy &CoLIN"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContenets 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "nlm"
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
Private Sub Form_Load()
datWords.DatabaseName = App.Path & "\nlp.mdb"
datWords.Refresh
End Sub
Public Sub cinWeb(htmlfile, Optional dr)
If IsMissing(dr) Then
    dr = "C:\"
End If
        Dim iret As Long
            iret = ShellExecute(Me.hwnd, _
            vbNullString, _
            htmlfile, _
            vbNullString, _
            dr, _
            SW_SHOWNORMAL)
End Sub
Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuCcomp_Click()
Clipboard.SetText Computer, 1
End Sub

Private Sub mnuChuman_Click()
Clipboard.SetText Human, 1

End Sub

Private Sub mnuClear_Click()
Set dbs = OpenDatabase(App.Path & "\nlp.mdb")
If MsgBox("Are you absolutely sure about this?", 276, "Delete All Data!") = vbNo Then Exit Sub
If MsgBox("This data is irrecoverable!", 257, "Proceed with delete?") = vbCancel Then Exit Sub

dbs.Execute "delete * from words"
datWords.Refresh
Computer = ""
Human = ""
End Sub

Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub mnuContenets_Click()
cinWeb App.Path & "\nlp.chm"
End Sub

Private Sub mnuPaste_Click()
Human = Clipboard.GetText
End Sub

Private Sub Talk_Click()
Dim words(20) As String
On Error Resume Next
Computer = ""
Human = LCase(Trim(Human))

Human = "^" & Human & "^"
Human = " " & Human & " "
Start = 1
cword = 1
For a = 1 To Len(Human)
    If Mid(Human, a, 1) = " " And a > 1 Then
        words(cword) = Mid(Human, Start + 1, a - Start)
        If Len(words(cword)) > Len(maxword) Then maxword = Trim(words(cword))

        Start = a
        cword = cword + 1
        
    End If
Next a
rf = "'adkevriy'"
For a = 1 To cword - 1

If Trim(words(a)) <> "" Then
datWords.Recordset.AddNew

middle = Trim(words(a))
rf = rf & ",'" & middle & "'"
'For b = 0 To 4
'    If optrate(b).Value = True Then Exit For
'Next b

If a > 1 Then Previous = Trim(words(a - 1))
If a < cword Then nextT = Trim(words(a + 1))
datWords.Recordset.FindFirst ("middle = '" & Trim(words(a)) & "' and previous = '" & Trim(words(a - 1)) & "' and next = '" & Trim(words(a + 1)) & "'")
If datWords.Recordset.NoMatch = True Then
    datWords.Recordset.AddNew
    datWords.Recordset.Fields(0) = middle
    datWords.Recordset.Fields(1) = Previous
    datWords.Recordset.Fields(2) = nextT
    datWords.Recordset.Update
    
End If
End If
Next a
Human = ""

Set dbs = OpenDatabase(App.Path & "\nlp.mdb")
Set tdf = dbs.OpenRecordset("select middle,count(middle) from words where middle in (" & rf & ") group by middle order by count(middle)")
maxword = tdf.Fields(0)

'
datWords.Recordset.MoveFirst
'
datWords.Recordset.FindFirst "middle = '" & maxword & "'"

Computer = maxword
wrd = choose(maxword, False)

Do While wrd <> ""


'
nextword = datWords.Recordset.Fields(1)
Computer = wrd & " " & Computer

wrd = choose(wrd, False)

'
num = (Int((datWords.Recordset.RecordCount * Rnd) + 1))
'
datWords.Recordset.MoveFirst
'
datWords.Recordset.Move num

'
datWords.Recordset.FindNext "middle = '" & nextword & "'"
'
If datWords.Recordset.NoMatch = True Then datWords.Recordset.FindLast "middle = '" & nextword & "'"



Loop
'
datWords.Recordset.MoveFirst
'
datWords.Recordset.FindFirst "middle = '" & maxword & "'"
wrd = choose(maxword, True)

Do While wrd <> ""

Computer = Computer & " " & wrd

wrd = choose(wrd, True)

'Do While datWords.Recordset.Fields(2) <> ""
'
nextword = datWords.Recordset.Fields(2)
'Computer = Computer & " " & nextword



'
num = (Int((datWords.Recordset.RecordCount * Rnd) + 1))
'
datWords.Recordset.MoveFirst
'
datWords.Recordset.Move num

'
datWords.Recordset.FindNext "middle = '" & nextword & "'"
'
If datWords.Recordset.NoMatch = True Then datWords.Recordset.FindLast "middle = '" & nextword & "'"


Loop
Computer = Mid(Computer, 2, Len(Computer) - 2)
comp = Computer
Computer = ""
picprog.ScaleWidth = Len(comp)
For c = 1 To Len(comp)
    Computer = Computer & Mid(comp, c, 1)
    For a = 1 To Int(50000 * Rnd) + 5000
        stuff = 5 * 5 * 5 * 5 * 5 * 5 * 2
    Next a
    Computer.Refresh
    pbsp.Width = c
    picprog.Refresh
Next c
Human.SetFocus
pbsp.Width = 0
End Sub
Public Function choose(word, forward As Boolean)
Set dbs = OpenDatabase(App.Path & "\nlp.mdb")
Set tdf = dbs.OpenRecordset("select * from words where middle = '" & word & "'")
rf = "'vkuseyvgwzelkbzwle'"
Do Until tdf.EOF
    chkstr = IIf(forward = True, tdf!Next, tdf!Previous)
    If InStr(1, Computer, chkstr) = 0 Or Len(rf) < 22 Then
        rf = rf & ",'" & IIf(forward = True, tdf!Next, tdf!Previous) & "'"
    
    End If
    tdf.MoveNext
Loop

    'New search routine - unstable and commented out for 1.1 bugfix - fixed 1.2 01/08/00
    Set rare = dbs.OpenRecordset("select middle,count(middle) from words where middle in (" & rf & ") and " & _
    IIf(forward = True, "previous = '" & word & "'", "next = '" & word & "'") & " group by middle order by count(middle)")
    'Set rare = dbs.OpenRecordset("select middle,count(middle) from words where middle in (" & rf & ") group by middle order by count(middle)")

If rare.EOF = True Then
    choose = ""
Else
    Do Until rare.EOF
        If Int((2 * Rnd) + 1) = 1 Then
            choose = rare!middle
            Exit Function
        Else
            rare.MoveNext
        End If
    Loop
    rare.MoveFirst
    choose = rare!middle
End If


End Function
