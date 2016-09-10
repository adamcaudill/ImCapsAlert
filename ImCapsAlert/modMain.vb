Module modMain

    Private Declare Auto Function PlaySound Lib "winmm.dll" _
    (ByVal lpszSoundName As String, ByVal hModule As Integer, _
    ByVal dwFlags As Integer) As Integer

    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short

    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Integer) As Short

    Private Const VK_CAPITAL As Short = &H14
    Private Const VK_LCONTROL As Short = &HA2

    Dim oThread As Threading.Thread

    Public Sub Main()
        oThread = New Threading.Thread(AddressOf MonitorCapsLock)
        oThread.Start()
    End Sub

    Private Function AppPath() As String
        Return System.AppDomain.CurrentDomain.BaseDirectory()
    End Function

    Private Sub MonitorCapsLock()
        Dim blnLastWasTrue As Boolean
        Do
            If CBool(GetKeyState(VK_CAPITAL)) = True Then
                'caps-lock is on
                If (CBool(GetAsyncKeyState(VK_LCONTROL)) = True) And (CBool(GetAsyncKeyState(VK_CAPITAL)) = True) Then
                    'control is down, time to exit
                    oThread.Abort()
                Else
                    'play the alarm
                    If blnLastWasTrue = False Then
                        blnLastWasTrue = True
                        PlayWav(False, AppPath() & "Woman Screem.wav")
                    End If
                End If
            Else
                blnLastWasTrue = False
            End If
            Threading.Thread.Sleep(750)
        Loop
    End Sub

    Private Function PlayWav(ByVal blnSystemDefault As Boolean, ByVal fileFullPath As String) As Boolean
        'return true if successful, false if otherwise
        Dim iRet As Integer = 0
        Const SND_FILENAME As Integer = &H20000
        Const SND_ASYNC As Integer = &H1
        Const SND_ALIAS As Integer = &H10000

        Try
            If blnSystemDefault = True Then
                iRet = PlaySound("MailBeep", 0, SND_ALIAS And SND_ASYNC)
            Else
                iRet = PlaySound(fileFullPath, 0, SND_FILENAME And SND_ASYNC)
            End If
        Catch
            'do nothing
        End Try

        Return CBool(iRet)
    End Function
End Module
