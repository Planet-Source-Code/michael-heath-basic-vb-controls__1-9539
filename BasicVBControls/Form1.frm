VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VB Basic Controls Example"
   ClientHeight    =   3165
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "ShellExecute API Example"
      Height          =   300
      Left            =   3825
      TabIndex        =   20
      Top             =   2520
      Width           =   3645
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse For Files"
      Height          =   315
      Left            =   3810
      TabIndex        =   19
      Top             =   2190
      Width           =   2355
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Color"
      Height          =   315
      Left            =   1410
      TabIndex        =   18
      Top             =   2820
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   30
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2820
      Width           =   1335
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1035
      LargeChange     =   4
      Left            =   90
      TabIndex        =   11
      Top             =   1320
      Width           =   225
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      LargeChange     =   10
      Left            =   90
      TabIndex        =   10
      Top             =   1110
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Background Color"
      Height          =   1965
      Left            =   5490
      TabIndex        =   6
      Top             =   150
      Width           =   1965
      Begin VB.OptionButton Option1 
         Caption         =   "Dither Red"
         Height          =   525
         Index           =   5
         Left            =   1080
         TabIndex        =   15
         Top             =   1290
         Width           =   825
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Dither Blue"
         Height          =   525
         Index           =   4
         Left            =   1080
         TabIndex        =   14
         Top             =   780
         Width           =   825
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Reset"
         Height          =   525
         Index           =   3
         Left            =   1080
         TabIndex        =   13
         Top             =   210
         Width           =   825
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Red"
         Height          =   525
         Index           =   2
         Left            =   60
         TabIndex        =   9
         Top             =   1290
         Width           =   1245
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Blue"
         Height          =   525
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   780
         Width           =   1245
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Desktop"
         Height          =   525
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   210
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Font"
      Height          =   1965
      Left            =   3810
      TabIndex        =   2
      Top             =   150
      Width           =   1665
      Begin VB.CheckBox Check1 
         Caption         =   "Underline"
         Height          =   525
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Top             =   1320
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Italic"
         Height          =   525
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   810
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bold"
         Height          =   525
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Width           =   1245
      End
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   90
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   315
      Left            =   6210
      TabIndex        =   0
      Top             =   2190
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "TextBox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   180
      TabIndex        =   17
      Top             =   2460
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   1380
      Width           =   480
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuHelpContact 
         Caption         =   "&Contact"
      End
      Begin VB.Menu mnuHelpBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Website"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click(Index As Integer)
' Set the style of font of the textbox
Select Case Index
    Case 0
        Text1.Font.Bold = (Check1(0).Value = vbChecked)
    Case 1
        Text1.Font.Italic = (Check1(1).Value = vbChecked)
    Case 2
        Text1.Font.Underline = (Check1(2).Value = vbChecked)
End Select
End Sub

Private Sub cmdBrowse_Click()
' Show the file browsing form
frmDrv.Show
End Sub

Private Sub cmdLaunch_Click()
MsgBox "This will launch a webpage from your Hard Drive.", vbOKOnly + vbInformation, "Basic VB Controls"
vLaunch App.Path & "\Default.htm"
End Sub

Private Sub Combo1_Click()
' Set the textbox's backcolor via the combo box
Select Case Combo1.Text
    Case "Blue"
        Text1.BackColor = vbBlue
    Case "Red"
        Text1.BackColor = vbRed
    Case "Desktop"
        Text1.BackColor = vbDesktop
    Case "White"
        Text1.BackColor = vbWhite
    Case Else
        MsgBox "That isn't a legitimate choice"
End Select

End Sub

Private Sub Command1_Click()
' Return system resources
Unload Me
Unload frmDrv
Set frmDrv = Nothing
Set Form1 = Nothing
End
End Sub

Private Sub Command2_Click()
' Error Handler
On Error GoTo CboErr
' Demonstrates Remove Item
Combo1.RemoveItem (0)
Exit Sub
CboErr:
    ' All the items are gone
    MsgBox "All Items in Combo Box have been removed", vbOKOnly + vbInformation, "Information"
End Sub

Private Sub Form_Load()
MsgBox "Running VB6 Basic Controls - By Michael Heath", vbOKOnly + vbInformation, "VB Basic Controls"
' Set the scroll bars max and min values
HScroll1.Min = 0
HScroll1.Max = 100
VScroll1.Min = 8
VScroll1.Max = 38
Label1.Caption = HScroll1.Value & " x " & VScroll1.Value
' Fill up the combo box
Combo1.AddItem "Blue"
Combo1.AddItem "Red"
Combo1.AddItem "Desktop"
Combo1.AddItem "White"
Combo1.AddItem "Something Else"
Combo1.Text = "Blue"
End Sub

Private Sub HScroll1_Change()
' Set label1's caption relative to the change in the scroll bars
Label1.Caption = HScroll1.Value & " x " & VScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
' Set label1's caption relative to the change in the scroll bars
Label1.Caption = HScroll1.Value & " x " & VScroll1.Value
End Sub

Private Sub mnuHelpAbout_Click()
MsgBox "VB Basic Controls v1.0 - By: Michael Heath", vbOKOnly + vbInformation, "About VB Basic Controls v1.0"
End Sub

Private Sub mnuHelpContact_Click()
vLaunch "mailto:mheath@morefreeware.com?Subject=BasicControls"
End Sub

Private Sub mnuHelpWeb_Click()
vLaunch App.Path & "\default.htm"
MsgBox "You are now being directed to my unfinished website.", vbOKOnly, "WebSite"
End Sub

Private Sub Option1_Click(Index As Integer)
' Set the form's backcolor according to the options the user selects
Select Case Index
    Case 0
        Form1.BackColor = vbDesktop
        Form1.Refresh
    Case 1
        Form1.BackColor = vbBlue
        Form1.Refresh
    Case 2
        Form1.BackColor = vbRed
        Form1.Refresh
    Case 3
        Form1.BackColor = vbButtonFace
        Form1.Refresh
    Case 4
        DitherBlue Me
    Case 5
        DitherRed Me
    
End Select
End Sub

Private Sub VScroll1_Change()
' Set label1's font size via the verticle scroll bar
Label1.Font.Size = VScroll1.Value
' Update label1's caption with the new values
Label1.Caption = HScroll1.Value & " x " & VScroll1.Value

End Sub

Private Sub VScroll1_Scroll()
' Set label1's font size via the verticle scroll bar
Label1.Font.Size = VScroll1.Value
' Update label1's caption with the new values
Label1.Caption = HScroll1.Value & " x " & VScroll1.Value
End Sub
