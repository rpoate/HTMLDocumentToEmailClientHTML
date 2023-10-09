Imports System.Text
Imports AngleSharp

Public Class InlineCSSToAttributes

    Private m_Document As Html.Dom.IHtmlDocument
    Private m_CssParser As InlineCssParser

    Public Sub New(Document As Html.Dom.IHtmlDocument)

        m_Document = Document
        m_CssParser = New InlineCssParser()

    End Sub

    Public Function ProcessHTML() As String

        For Each oImg In m_Document.GetElementsByTagName("img")
            ProcessElement(oImg, {"width", "height"}, "px")
        Next

        For Each oTable In m_Document.GetElementsByTagName("table")
            ProcessElement(oTable, {"width", "height"}, Nothing)
        Next

        Return m_Document.ToHtml

    End Function

    Private Sub ProcessElement(ByRef Element As Html.Dom.IHtmlElement, CSSAttributesToSimplfy() As String, removeString As String)

        For Each sAttr In CSSAttributesToSimplfy
            m_CssParser.InlineCss = Element.GetAttribute("style")
            If Not String.IsNullOrEmpty(removeString) Then
                Try
                    Element.SetAttribute(sAttr, m_CssParser.GetValue(sAttr).Replace(removeString, ""))
                Catch

                End Try
            Else
                Element.SetAttribute(sAttr, m_CssParser.GetValue(sAttr))
            End If
        Next

    End Sub

    Private Class InlineCssParser

        Dim lDic As Dictionary(Of String, String)

        Public Property InlineCss As String
            Get
                Return InlineCss
            End Get
            Set(value As String)
                lDic = parse(value)
            End Set
        End Property

        Public Sub SetValue(ByVal name As String, ByVal value As String)
            setDictionary(lDic, name, value)
            InlineCss = combine(lDic)
        End Sub

        Public Function GetValue(ByVal name As String) As String
            Return getDictionary(name)
        End Function

        Public Sub RemoveItem(ByVal name As String)
            lDic = parse(InlineCss)
            removeFromDictionary(name)
            InlineCss = combine(lDic)
        End Sub

        Public Function HasName(ByVal name As String) As Boolean
            Return hasDictionary(name)
        End Function

        Private Function combine(ByVal dic As ICollection(Of KeyValuePair(Of String, String))) As String
            If dic Is Nothing OrElse dic.Count <= 0 Then
                Return String.Empty
            Else
                Dim result = New StringBuilder()

                For Each pair In dic
                    If result.Length > 0 Then result.Append("; ")
                    result.AppendFormat("{0}: {1}", pair.Key, pair.Value)
                Next

                Return result.ToString()
            End If
        End Function

        Private Function parse(ByVal raw As String) As Dictionary(Of String, String)
            Dim result = New Dictionary(Of String, String)()

            If raw IsNot Nothing Then
                Dim firsts = raw.Split({";"c}, StringSplitOptions.RemoveEmptyEntries)

                For Each first In firsts
                    Dim f = first.Trim()

                    If Not String.IsNullOrEmpty(f) Then
                        Dim seconds = f.Split({":"c}, StringSplitOptions.RemoveEmptyEntries)

                        If seconds.Length = 2 Then
                            Dim a = seconds(0).Trim().ToLowerInvariant()
                            Dim b = seconds(1).Trim()

                            If Not String.IsNullOrEmpty(a) AndAlso Not String.IsNullOrEmpty(b) Then
                                setDictionary(result, a, b)
                            End If
                        End If
                    End If
                Next
            End If

            Return result
        End Function

        Private Function getDictionary(ByVal a As String) As String
            Return If(lDic.ContainsKey(a.ToLowerInvariant()), lDic(a.ToLowerInvariant()), Nothing)
        End Function

        Private Function hasDictionary(ByVal a As String) As Boolean
            Return lDic.ContainsKey(a.ToLowerInvariant())
        End Function

        Private Sub removeFromDictionary(ByVal a As String)
            If lDic.ContainsKey(a.ToLowerInvariant()) Then
                lDic.Remove(a.ToLowerInvariant())
            End If
        End Sub

        Private Sub setDictionary(ByRef dic As Dictionary(Of String, String), ByVal a As String, ByVal b As String)
            If dic.ContainsKey(a.ToLowerInvariant()) Then
                dic(a.ToLowerInvariant()) = b
            Else
                dic.Add(a.ToLowerInvariant(), b)
            End If
        End Sub
    End Class
End Class

