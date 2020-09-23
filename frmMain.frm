VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Legal Notice Manager"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCaption 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.CheckBox chkShowOnce 
      Caption         =   "Show Legal Notice Only Once"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Form"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Caption:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
  txtCaption.Text = ""
  txtMessage.Text = ""
End Sub

Private Sub cmdDelete_Click()
  cmdClear_Click
  SetLegalNotice "", "", True
  StartWithWindows ApplicationName, "", "", True
End Sub

Private Sub cmdSave_Click()
  Dim Caption As String
  Dim Msg As String
  
  Caption = Trim(txtCaption.Text)
  Msg = Trim(txtMessage.Text)
  
  SetLegalNotice Caption, Msg
  
  If chkShowOnce.Value = 1 Then
    StartWithWindows ApplicationName, App.Path & "\" & App.EXEName, SHOWONCE
  Else
    StartWithWindows ApplicationName, "", "", True
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    End
  End If
End Sub

Private Sub Form_Load()
  Dim WindowsType As String
  
  If IsWin95 Or IsWin98 Or IsWinME Then 'we're on a system running one of the WIN9x kernels
    WindowsType = "Windows"
  ElseIf IsWinNT Or IsWin2K Or IsWinXP Then 'we're on a system running a WINNT kernel
    WindowsType = "Windows NT"
  End If

  txtCaption.Text = GETSTRING(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\" & WindowsType & "\CurrentVersion\Winlogon", "LegalNoticeCaption")
  txtMessage.Text = GETSTRING(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\" & WindowsType & "\CurrentVersion\Winlogon", "LegalNoticeText")

End Sub
