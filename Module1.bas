Attribute VB_Name = "Module1"
'Coded By Rajendra Khope, Pune, India
'App Name: YoutubeVideoDownloader
'Use: Software to Serach and Download YouTube Video.
'
'For more Info: http://youtube.com
'
'Email : bkrajendra@gmail.com
'Web: http://www.figmentsol.com/ytvd/

Option Explicit

Public Type VideoInfo
    vID As String
    vTitle As String
    vDuration As String
    vAuthor As String
    'vLength As String
    vThumbURL As String
    vDesciption As String
    vCategory As String
    vViews As String
    vPub As String
    vRatings As String
    'more to come
End Type
Public vInfo As VideoInfo
Dim FileSize_Current As Long
'Drag Form
Private Declare Sub ReleaseCapture Lib "User32" ()
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long
'Drag Form
Public Sub FormDrag(theform As Form)
    ReleaseCapture
    Call SendMessage(theform.hWnd, &HA1, 2, 0&)
End Sub

Function JamesBond(Text As String, Pattern As String) As String
'Coded By Rajendra Khope
'This is a regular Expression in vb
'include a reference to "Microsoft VBScript Regular Expressions" in your project
'You can find more info @
'http://www.regular-expressions.info/vb.html
    Dim Regex As RegExp
    Dim Matches As Variant

    Set Regex = New RegExp
    Regex.Pattern = Pattern
    Set Matches = Regex.Execute(Text)
    If Matches.Count = 0 Then
        JamesBond = ""
        Exit Function
    End If

    JamesBond = Matches(0).SubMatches(0)
End Function
'Testing.......GetAnyThing
Public Function GetAnyThing(strResp, strT As String, strE As String) As String
    On Error Resume Next
    Dim pos1 As Long
    Dim pos2 As Long
    
    pos1 = InStr(1, strResp, strT) + Len(strT)
    pos2 = InStr(pos1, strResp, strE)
    GetAnyThing = Mid(strResp, pos1, pos2 - pos1)
End Function
Function UrlParser(strGarbage As String) As String
    UrlParser = JamesBond(strGarbage, Chr(34) & "fmt_url_map" & Chr(34) & ": " & Chr(34) & "([^""]+)")
End Function
Function URLDecode(strEncUrl)
Dim I As Integer
Dim st, sR As String
        strEncUrl = Replace(strEncUrl, "+", " ")
        For I = 1 To Len(strEncUrl)
            st = Mid(strEncUrl, I, 1)
            If st = "%" Then
                If I + 2 < Len(strEncUrl) Then
                    sR = sR & _
                        Chr(CLng("&H" & Mid(strEncUrl, I + 1, 2)))
                    I = I + 2
                End If
            Else
                sR = sR & st
            End If
        Next
        URLDecode = sR
End Function
Function DWFile(dwLik As String, FileName As String, proGress As ProgressBar, centProgress As Label, ctlInet As Inet)
On Error GoTo ErrorControl
Dim FileSize As Long
Dim sz As Double
Dim FileRemaining As Long
Dim FileNumber As Integer
Dim FileData() As Byte
'Dim FileSize_Current As Long
Dim PBValue As Integer
    frmDW.Timer1.Enabled = True
    ctlInet.Execute Trim(dwLik), "GET"
    Do While ctlInet.StillExecuting
        DoEvents
    Loop
    
    'Retrieve file size from content header
    'You can refer this Link for this:
    'http://support.microsoft.com/kb/163653
    FileSize = ctlInet.GetHeader("Content-Length")
        sz = FileSize / 1000
        'lblFSize.Caption = sz & " Kb"
    FileRemaining = FileSize
    FileSize_Current = 0
   
    FileNumber = FreeFile
    Open App.Path & "\" & FileName For Binary Access Write As #FileNumber
    
    Do Until FileRemaining = 0
        If FileRemaining > 1024 Then
            FileData = ctlInet.GetChunk(1024, icByteArray)
            FileRemaining = FileRemaining - 1024
        Else
            FileData = ctlInet.GetChunk(FileRemaining, icByteArray)
            FileRemaining = 0
        End If
        
        FileSize_Current = FileSize - FileRemaining
        PBValue = CInt((100 / FileSize) * FileSize_Current)
        centProgress.Caption = PBValue & " % "
        proGress.Value = PBValue ' * 40
        Put #FileNumber, , FileData
    Loop
    
    Close #FileNumber
    frmDW.Timer1.Enabled = False
    Exit Function
ErrorControl:
    MsgBox "Error-" & Err.Description
End Function
Function GetAllInfo(strResponse As String)
    Dim strTempHolder As String
    
    strTempHolder = GetAnyThing(strResponse, "VIDEO_TITLE': '", "',")
    strTempHolder = Replace(strTempHolder, "/", "")
    strTempHolder = Replace(strTempHolder, "\", "")
    strTempHolder = Replace(strTempHolder, "'", "")
    strTempHolder = Replace(strTempHolder, ":", "")
    strTempHolder = Replace(strTempHolder, "(", "")
    strTempHolder = Replace(strTempHolder, ")", "")
    
    vInfo.vTitle = strTempHolder
    
    strTempHolder = GetAnyThing(strResponse, "data-discoverbox-username=""", """ >")
    vInfo.vAuthor = strTempHolder
    
    strTempHolder = GetAnyThing(strResponse, "'VIDEO_ID': '", "'")
    vInfo.vID = strTempHolder
    
    strTempHolder = GetAnyThing(strResponse, "watch-video-added post-date"">", "</span>")
    vInfo.vPub = strTempHolder
    
    strTempHolder = GetAnyThing(strResponse, "VideoCategoryLink');"">", "</a>")
    vInfo.vCategory = strTempHolder
    
    strTempHolder = GetAnyThing(strResponse, "watch-video-desc description"">", "</span>")
    vInfo.vDesciption = Mid(strTempHolder, InStr(1, strTempHolder, "<span >") + 7, Len(strTempHolder))
    
    strTempHolder = GetAnyThing(strResponse, "rv.2.thumbnailUrl"": """, """,")
    vInfo.vThumbURL = URLDecode(strTempHolder)
    'watch-view-count">
    
    strTempHolder = GetAnyThing(strResponse, "length_seconds"": """, """,")
    vInfo.vDuration = URLDecode(strTempHolder)
    
    strTempHolder = GetAnyThing(strResponse, "master-sprite ratingL ratingL-", """")
    vInfo.vRatings = URLDecode(strTempHolder)
    
    strTempHolder = GetAnyThing(strResponse, "watch-view-count"">", "</span>")
    vInfo.vViews = URLDecode(strTempHolder)
    
End Function
Public Sub SearchEngine(strPageHTML As String, lstTitle As ListBox, lstID As ListBox)
'Coded By Rajendra Khope
'Searches for query
'On Error GoTo errr

Dim strTemp, strTitle, strID As String
Dim intr1, intr2, pointer, I As Long
'Dim ResultsCounter As Integer
        lstTitle.Clear
        lstID.Clear

    I = 0
    pointer = 1
    
    For I = 1 To Len(strPageHTML)

        intr1 = InStr(pointer, strPageHTML, "video-entry yt-uix-hovercard", vbTextCompare)
        pointer = intr1 + 1
        If intr1 = 0 Then
            GoTo comeout
        End If
        
        intr2 = InStr(pointer, strPageHTML, "video-cell", vbTextCompare)
        
        If intr2 = 0 Then
            GoTo comeout
        End If
        
    strTemp = Mid(strPageHTML, pointer, intr2 - intr1)
    
    strTitle = GetAnyThing(strTemp, "hovercard-title"" >", "</strong>")
    strID = GetAnyThing(strTemp, "watch?v=", """")
    'Add Video Title
    If strTitle <> "" Then
        lstTitle.AddItem strTitle
        lstID.AddItem strID
    End If
    
    Next

comeout:
Exit Sub
errr:

MsgBox "Error: " & Err.Description
'dbug "Error: " & vbCrLf & Err.Description & " Try Again..!"
End Sub
Function HTMLCleaner(Html As String)
    Dim TotalLength, I, CurPosition As Integer
    Dim OneChara, PureText As String
    
    TotalLength = Len(Html)
    
    For I = 1 To TotalLength
        OneChara = Mid(Html, I, 1)
        CurPosition = I
        If (OneChara = "<") Then
            CurPosition = I
            Do While (Mid(Html, CurPosition, 1) <> ">")
                CurPosition = CurPosition + 1
            Loop
            PureText = PureText & " "
        Else
            PureText = PureText & OneChara
        End If
        I = CurPosition
    Next
    HTMLCleaner = PureText
End Function
'Not used...
Function GetResultEntryInfo(YouTube_Hovercard As String)
    Dim strTempHolder As String
    
    strTempHolder = GetAnyThing(YouTube_Hovercard, "watch?v=", """")
    vInfo.vID = strTempHolder
    
    strTempHolder = GetAnyThing(YouTube_Hovercard, "hovercard-title"" >", "</strong>")
    vInfo.vTitle = strTempHolder
    
    strTempHolder = GetAnyThing(YouTube_Hovercard, "hovercard-username"">", "</p>")
    vInfo.vAuthor = strTempHolder
    
    strTempHolder = GetAnyThing(YouTube_Hovercard, "hovercard-category"">", "</p>")
    vInfo.vCategory = strTempHolder
    
    strTempHolder = GetAnyThing(YouTube_Hovercard, "hovercard-description"" >", "<div>")
    vInfo.vDesciption = HTMLCleaner(strTempHolder)
    
    strTempHolder = GetAnyThing(YouTube_Hovercard, " thumb=""", """")
    vInfo.vThumbURL = strTempHolder
    
    strTempHolder = GetAnyThing(YouTube_Hovercard, "hovercard-views"">", "</span>")
    vInfo.vThumbURL = strTempHolder
End Function
Public Function GetTitleImage(Connector As Inet, imageURL As String, FileName As String) As Boolean
 On Error GoTo errorHandler
    Dim ImageData() As Byte
    Dim FiePath As String
    
    If Connector.StillExecuting = True Then Exit Function
    ImageData() = Connector.OpenURL(imageURL, icByteArray)
    
    Do While Connector.StillExecuting
        DoEvents
    Loop
   
    FiePath = App.Path & "\Cache\" + FileName
    Open FiePath For Binary Access Write As #1
    Put #1, , ImageData()
    Close #1
    
    GetTitleImage = True
    Exit Function
errorHandler:
GetTitleImage = False
End Function
