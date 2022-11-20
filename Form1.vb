Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        HtmlEditControl1.CSSText = "body {font-family: Arial} p {font-size: 12pt} a {text-decoration: underline}"

        Dim oButton = HtmlEditControl1.ToolStripItems.Add("Extract HTML For Email Client")

        oButton.Padding = New Padding(3)
        AddHandler oButton.Click, AddressOf oButton_CLick

    End Sub

    Private Sub oButton_CLick(sender As Object, e As EventArgs)

        Dim sCurrentHTML = "<html><head><style>" & HtmlEditControl1.CSSText & "</style></head><body>" & HtmlEditControl1.DocumentHTML & "</body></html>"
        Dim preMail = New PreMailer.Net.PreMailer(sCurrentHTML)

        preMail.MoveCssInline()

        Dim StylesToAttr As New InlineCSSToAttributes(preMail.Document)

        MsgBox(StylesToAttr.ProcessHTML)
    End Sub
End Class
