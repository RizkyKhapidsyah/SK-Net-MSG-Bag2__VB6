VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Messenger "
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMsg 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4815
   End
   Begin VB.TextBox txtServer 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Message"
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Computer"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'Form1.MousePointer = 11 'vbHourglass
Dim wCommand As String
wCommand = "c:\winnt\system32\net send "
wCommand = wCommand & Trim(txtServer) & " "
wCommand = wCommand & Trim(txtMsg)
Shell (wCommand)
txtMsg = ""
txtMsg.SetFocus
'Form1.MousePointer = 0 'vbNormal
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub txtMsg_GotFocus()
txtMsg.SelStart = 0
txtMsg.SelLength = Len(txtMsg)
End Sub

Private Sub txtMsg_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1.SetFocus
End Sub

Private Sub txtServer_GotFocus()
txtServer.SelStart = 0
txtServer.SelLength = Len(txtServer)
End Sub

Private Sub txtServer_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtMsg.SetFocus
End Sub
