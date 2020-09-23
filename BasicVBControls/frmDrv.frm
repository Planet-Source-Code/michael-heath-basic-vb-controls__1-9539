VERSION 5.00
Begin VB.Form frmDrv 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Browse For Files"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7395
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox lstPath 
      Height          =   345
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3000
      Width           =   7185
   End
   Begin VB.DriveListBox Drv1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   2895
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   3180
      TabIndex        =   1
      Top             =   150
      Width           =   4065
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   90
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "frmDrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
' Set the files to view in the File1 box
File1.Path = Dir1.Path
End Sub

Private Sub Drv1_Change()
' Set the Dir1 from the drives you're on
Dir1.Path = Drv1.Drive
End Sub

Private Sub File1_Click()
' Place the full path and filename in the textbox and
' in the caption of the form
lstPath.Text = Dir1.Path & File1.FileName
Me.Caption = Dir1.Path & File1.FileName
End Sub
