Attribute VB_Name = "modMain"
Public Const ApplicationName = "LegalNoticeMangager"
Public Const SHOWONCE = "SHOW_ONCE"

Sub Main()
'Here we go, everybody ready?  :-)

  Dim CmdLine As String
  
  CmdLine = Trim(UCase(Command$()))
  
  Select Case CmdLine
    Case SHOWONCE 'the application has been started from the registry (most likely anyway)
      'the user elected to have it show once so we're going to delete
      '   the entries in the registry and take this application out of
      '   the "run at startup" line up.
      SetLegalNotice "", "", True
      StartWithWindows ApplicationName, "", "", True
    Case "" 'the application has been started from Windows Explorer
      'the user wants to set/save the Legal Notice parameters so we're going to
      '   show them the form we built to do just that.
      frmMain.Show
  End Select
  
End Sub

Public Function SetLegalNotice(ByVal Caption As String, ByVal Message As String, Optional ByVal Remove As Boolean = False) As Boolean
'accepts the caption and the message to display in the legal notice message box that fires before the
'   rest of windows loads.  Also accepts an optional parameter to remove said legal notice.
  
  Dim WindowsType As String
  
  If IsWin95 Or IsWin98 Or IsWinME Then 'we're on a system running one of the WIN9x kernels
    WindowsType = "Windows"
  ElseIf IsWinNT Or IsWin2K Or IsWinXP Then 'we're on a system running a WINNT kernel
    WindowsType = "Windows NT"
  End If
  
  If Remove Then
    Caption = ""
    Message = ""
  End If
  
  SAVESTRING HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\" & WindowsType & "\CurrentVersion\Winlogon", "LegalNoticeCaption", Caption
  SAVESTRING HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\" & WindowsType & "\CurrentVersion\Winlogon", "LegalNoticeText", Message
  
End Function

Public Function StartWithWindows(ByVal AppName As String, ByVal AppPath As String, ByVal CmdLineArg As String, Optional ByVal Remove As Boolean = False) As Boolean
'accepts the full path to an exe and its name and adds that application to the startup lineup to start with
'   windows next time it loads.  Also accepts a command line argument to add to the application
'   path along with an optional parameter to remove that particular application from the registry \Run\
'   key.
  
  If Remove = False Then 'they want to add the application
    SAVESTRING HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", AppName, Trim(AppPath) & " " & Trim(CmdLineArg)
  Else 'they want to remove the application
    DeleteRegistryValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", AppName
  End If
  
End Function

Public Function IsWin98() As Boolean
'obvious

  Dim ProductName As String
  IsWin98 = False
  
  ProductName = GETSTRING(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProductName")
  ProductName = Trim(ProductName)
  
  If ProductName = "Microsoft Windows 98" Then
    IsWin98 = True
  End If

End Function

Public Function IsWin95() As Boolean
'obvious

  Dim ProductName As String
  IsWin95 = False
  
  ProductName = GETSTRING(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProductName")
  ProductName = Trim(ProductName)
  
  If ProductName = "Microsoft Windows 95" Then
    IsWin95 = True
  End If
End Function

Public Function IsWinNT() As Boolean
'obvious

  Dim ProductName As String
  IsWinNT = False
  
  ProductName = GETSTRING(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProductName")
  ProductName = Trim(ProductName)
  
  If ProductName = "Microsoft Windows NT" Then
    IsWinNT = True
  End If
End Function

Public Function IsWin2K() As Boolean
'obvious

  Dim ProductName As String
  IsWin2K = False
  
  ProductName = GETSTRING(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "ProductName")
  ProductName = Trim(ProductName)
  
  If ProductName = "Microsoft Windows 2000" Then
    IsWin2K = True
  End If
End Function

Public Function IsWinXP() As Boolean
'obvious

  Dim ProductName As String
  IsWinXP = False
  
  ProductName = GETSTRING(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "ProductName")
  ProductName = Trim(ProductName)
  
  If ProductName = "Microsoft Windows XP" Then
    IsWinXP = True
  End If
End Function

Public Function IsWinME() As Boolean
'obvious

  Dim ProductName As String
  IsWinME = False
  
  ProductName = GETSTRING(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProductName")
  ProductName = Trim(ProductName)
  
  If ProductName = "Microsoft Windows ME" Then
    IsWinME = True
  End If
End Function

Public Function StripTerminator(ByVal strString As String) As String
'This function accepts a string as its only argument.
'It then strips off the null (chr(0)) off the end.
'More specifically, it strips everything to the left of
'the first null including the null and returns everything
'to the right of the null.

    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

