Public Class Form1
    'Dim stxt As String = Strings.Left(Clipboard.GetText, 200)
    'Dim stxt As String = Replace(Replace(Replace(Clipboard.GetText, " ", "%20"), Chr(13) & Chr(10), ""), Chr(9), "%20")
    Private Sub TextBox1_Click(sender As Object, e As System.EventArgs) Handles TextBox1.Click
        'TextBox1.Text = stxt
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged
        'If TextBox1.Text <> "" Then
        '    查詢百度()
        '    查詢百度百科()
        '    查詢Yahoo()
        '    查詢Bing()
        '    查詢Google()
        '    'Me.TextBox1.Copy()
        '    'Clipboard.SetText(Replace(TextBox1.Text, "%20", " "))
        '    Clipboard.SetText(TextBox1.Text)
        '    'If dbCompactDatabase Then 壓縮檔案("codes.mdb")
        '    End
        'End If

    End Sub

    Private Sub Form1_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        'TextBox1.Text = stxt
    End Sub

    Private Sub Form1_Click(sender As Object, e As System.EventArgs) Handles Me.Click
        'TextBox1.Text = stxt
    End Sub

    Private Sub Form1_GotFocus(sender As Object, e As System.EventArgs) Handles Me.GotFocus
        'TextBox1.Text = stxt
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        'If e = System.ConsoleKey.Escape Then

        'End If
    End Sub


    Public Sub New()
        ' 此為設計工具所需的呼叫。
        InitializeComponent()

        ' 在 InitializeComponent() 呼叫之後加入任何初始設定。

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        'UTF8()
        '查詢字串轉換_Big5碼("春")
        End
    End Sub
End Class
