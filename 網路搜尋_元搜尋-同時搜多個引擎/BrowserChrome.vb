Imports System.IO
Imports Microsoft.Win32
Public Class BrowserChrome

    'https://gist.github.com/fredrikhaglund/43aea7522f9e844d3e7b
    Private Const ChromeAppKey As String =
            "\Software\Microsoft\Windows\CurrentVersion\App Paths\
                    chrome.exe"
    Public ReadOnly Property ChromeAppFileName As String
        Get
            Dim caf As String = Registry.GetValue("HKEY_LOCAL_MACHINE" &
                ChromeAppKey, "", Nothing)
            ChromeAppFileName = IIf(caf Is Nothing,
            Registry.GetValue("HKEY_CURRENT_USER" + ChromeAppKey,
                    "", Nothing), caf) 'https://bit.ly/3gxttQR
            If ChromeAppFileName Is Nothing Then
                Const chromeFullname = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
                If File.Exists(chromeFullname) Then Return chromeFullname
            End If
        End Get
    End Property

    Public Sub OpenLinkChrome(url As String)
        Dim crmAppFileName As String = ChromeAppFileName
        If String.IsNullOrEmpty(crmAppFileName) Then 'https://bit.ly/3dOuemY
            Throw New Exception("Could not find chrome.exe!")
        End If
        Process.Start(ChromeAppFileName, urlRegx(url))
    End Sub

    Private ReadOnly Property urlRegx(url As String) As String
        Get
            Dim replWds() As String = {"""", "%22", " ", "%20"} ' "http//", ""}
            Dim i As Byte
            Do While (i < replWds.Length)
                url = url.Replace(replWds(i), replWds(i + 1))
                i += 2
            Loop
            urlRegx = url
        End Get
    End Property
End Class
