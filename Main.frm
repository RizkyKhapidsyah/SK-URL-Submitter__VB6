VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form MikroSubmit 
   BackColor       =   &H00000000&
   Caption         =   "MikroSubmit"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Main.frx":000C
   ScaleHeight     =   4005
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox SubmitBox 
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   4680
      Width           =   7335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   3480
      Width           =   2535
   End
   Begin SHDocVwCtl.WebBrowser MikroWeb 
      Height          =   1725
      Left            =   5160
      TabIndex        =   5
      Top             =   2340
      Width           =   2415
      ExtentX         =   4260
      ExtentY         =   3043
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton CmdGoogle 
      Caption         =   "Submit To Google"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox email 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   4695
   End
   Begin VB.TextBox desc 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   4695
   End
   Begin VB.TextBox url 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Text            =   "http://"
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox site 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "MikroSubmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WebUrl
Dim Verify
Dim Name1
Dim SiteName
Dim Desciption

Private Sub CmdGoogle_Click()
SiteName = site.Text
WebUrl = url.Text
Call validate
If Verify = 1 Then
MsgBox "Please fill out all the fields"
Exit Sub
End If
SubmitBox.Text = "http://www.google.com/addurl?q=" & WebUrl & "&dq=" & SiteName & "&I1.x=37&I1.y=15"
MikroWeb.Navigate (SubmitBox.Text)

End Sub

Sub validate()
If url.Text = "http://" Then
End If
If site.Text = "" Then
End If
If desc.Text = "" Then
End If
If email.Text = "" Then
Verify = 1
Else
Verify = 0
MikroWeb.Visible = True
Command1.Visible = True
End If
End Sub

Private Sub Command1_Click()
MikroWeb.Visible = False
Command1.Visible = False
End Sub

Private Sub Form_Load()
MikroWeb.Visible = False
Command1.Visible = False
End Sub
