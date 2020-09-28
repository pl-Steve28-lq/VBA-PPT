Private Sub WeatherLoad_Click()
    Dim WinHttp As New WinHttpRequest
    Dim Res As String
    Dim Time As String

    WinHttp.Open "Get", "https://search.naver.com/search.naver?sm=top_hty&fbm=1&ie=utf8&query=" + Slide1.Region.Text + "%EB%82%A0%EC%94%A8"

    WinHttp.SetRequestHeader "User-Agent", "Mozilla/4.0(compatible;MSIE 6.0; Windows NT 5.0)"

    WinHttp.Send
    
    Res = WinHttp.ResponseText
    
    Weather.Text = Between(Res, "<p class=""cast_txt"">", "</p>")
    Temp.Text = Between(Res, "<span class=""todaytemp"">", "</span>")
End Sub

Function Between(str As String, a As String, b As String) As String
    On Error Resume Next
    Between = Split(Split(str, a)(1), b)(0)
    If Err.Number <> 0 Then
        Between = "처리 오류"
        Err.Clear
    End If
End Function

Private Sub Clear_Click()
    Weather.Text = ""
    Temp.Text = ""
    Region.Text = ""
End Sub
