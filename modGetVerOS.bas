Attribute VB_Name = "modGetVerOS"
'
'Detecting OS (Unicode) Windows 95 - Windows 7, pappsegull@yahoo.se nov 5 2010
'
Option Explicit
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExW" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion      As Long
   dwMinorVersion      As Long
   dwBuildNumber       As Long
   dwPlatformId        As Long
   szCSDVersion        As String * 128 'Maintenance string for PSS usage
End Type
Public Enum VerOS
    [Not Detected]
    [Windows 95]
    [Windows 98]
    [Windows ME]
    [Windows NT 3.51]
    [Windows NT 4.0]
    [Windows 2000]
    [Windows XP]
    [Windows 2003]
    [Windows Vista]
    [Windows 7]
End Enum

Public Function GetVerOS(Optional RetInfo$) As VerOS
'                Win95 Win98 WinME NT 3.51 NT 4.0 Win2000 WinXP Win2003 Vista  Win7
' ------------------------------------------------------------------------------------
'dwPlatFormID      1     1     1      2      2      2       2     2       2      2
'dwMajorVersion    4     4     4      3      4      5       5     5       6      6
'dwMinorVersion    0    10    90     51      0      0       1     2       0      1

Dim s$, OS As OSVERSIONINFO
    With OS
        .dwOSVersionInfoSize = LenB(OS): If Not CBool(GetVersionEx(OS)) Then Exit Function
        Select Case .dwPlatformId
            Case 1 '< NT 3.51
                Select Case .dwMajorVersion
                    Case 4
                        Select Case .dwMinorVersion
                            Case 0: GetVerOS = [Windows 95]
                            Case 10: GetVerOS = [Windows 98]
                            Case 90: GetVerOS = [Windows ME]
                            Case Else: GetVerOS = [Not Detected]
                        End Select
                    Case Else: GetVerOS = [Not Detected]
                End Select
            Case 2
                Select Case .dwMajorVersion
                    Case 3: GetVerOS = [Windows NT 3.51]
                    Case 4: GetVerOS = [Windows NT 4.0]
                    Case 5 '< Vista
                        Select Case .dwMinorVersion
                            Case 0: GetVerOS = [Windows 2000]
                            Case 1: GetVerOS = [Windows XP]
                            Case 2: GetVerOS = [Windows 2003]
                            Case Else: GetVerOS = [Not Detected]
                        End Select
                    Case 6
                        Select Case .dwMinorVersion
                            Case 0: GetVerOS = [Windows Vista]
                            Case 1: GetVerOS = [Windows 7]
                            Case Else: GetVerOS = [Not Detected]
                        End Select
                    Case Else: GetVerOS = [Not Detected]
                End Select
        End Select
        s$ = StrConv(.szCSDVersion, vbFromUnicode): If GetVerOS > 0 Then RetInfo$ = "Windows "
        RetInfo$ = "Your OS is " & RetInfo$ & Choose(GetVerOS + 1, "not detected", _
          "95", "98", "ME", "NT 3.51", "NT 4.0", "2000", "XP", "2003", "Vista", "7") & _
          ", Ver: " & .dwPlatformId & "." & .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber & _
          "  (" & Left$(s$, InStr(s$, vbNullChar) - 1) & ")."
    End With
End Function

Sub Main()
Dim s$ 'Just a demo of the module.
    If GetVerOS(s$) <> [Not Detected] Then MsgBox s$, vbInformation
    End
End Sub
