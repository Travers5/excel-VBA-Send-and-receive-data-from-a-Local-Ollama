Attribute VB_Name = "OllamaSpillModule"
Option Explicit

Public Const OLLAMA_BASE_URL As String = "http://127.0.0.1:11434"
Private Const MAX_CELL_CHARS As Long = 32767
Private Const SAFE_CELL_CHARS As Long = 32000

Public Sub ShowOllamaPromptGui()
    frmOllamaPrompt.Show
End Sub

Public Function GetPromptFromRanges(ByVal rangeExpression As String) As String
    Dim cleaned As String
    cleaned = Replace(rangeExpression, vbCr, vbNullString)
    cleaned = Replace(cleaned, vbLf, vbNullString)

    Dim parts() As String
    parts = Split(cleaned, ",")

    Dim i As Long
    Dim output As String
    Dim ws As Worksheet
    Set ws = ActiveSheet

    For i = LBound(parts) To UBound(parts)
        Dim token As String
        token = Trim$(parts(i))
        If Len(token) > 0 Then
            Dim rng As Range
            On Error GoTo InvalidRange
            Set rng = ws.Range(token)
            On Error GoTo 0

            Dim c As Range
            For Each c In rng.Cells
                If Len(CStr(c.Value2)) > 0 Then
                    If Len(output) > 0 Then
                        output = output & vbLf
                    End If
                    output = output & CStr(c.Value2)
                End If
            Next c
        End If
    Next i

    GetPromptFromRanges = output
    Exit Function

InvalidRange:
    Err.Raise vbObjectError + 1000, "GetPromptFromRanges", _
              "Invalid range token: '" & token & "'. Use comma-separated addresses (e.g., A1,A2:B3)."
End Function

Public Function GetAvailableModels() As Collection
    Dim models As New Collection
    Dim body As String
    body = HttpGet(OLLAMA_BASE_URL & "/api/tags")

    Dim matches As Object
    Set matches = RegexFindAll(body, "\"name\"\s*:\s*\"([^\"]+)\"")

    Dim i As Long
    For i = 0 To matches.Count - 1
        models.Add matches(i).SubMatches(0)
    Next i

    If models.Count = 0 Then
        models.Add "llama3"
    End If

    Set GetAvailableModels = models
End Function

Public Function GenerateFromOllama(ByVal modelName As String, ByVal promptText As String) As String
    If Len(Trim$(modelName)) = 0 Then
        Err.Raise vbObjectError + 1100, "GenerateFromOllama", "Model name is required."
    End If
    If Len(promptText) = 0 Then
        Err.Raise vbObjectError + 1101, "GenerateFromOllama", "Prompt cannot be empty."
    End If

    Dim payload As String
    payload = "{""model"":""" & JsonEscape(modelName) & """,""prompt"":""" & JsonEscape(promptText) & """,""stream"":false}"

    Dim body As String
    body = HttpPostJson(OLLAMA_BASE_URL & "/api/generate", payload)

    Dim matches As Object
    Set matches = RegexFindAll(body, "\"response\"\s*:\s*\"((?:\\.|[^\"])*)\"")
    If matches.Count = 0 Then
        Err.Raise vbObjectError + 1102, "GenerateFromOllama", "No response field found in Ollama output."
    End If

    GenerateFromOllama = JsonUnescape(matches(0).SubMatches(0))
End Function

Public Sub WriteSpilledText(ByVal anchorAddress As String, ByVal fullText As String)
    Dim startCell As Range
    Set startCell = ActiveSheet.Range(anchorAddress)

    Dim chunks As Collection
    Set chunks = SplitIntoCellSizedChunks(fullText, SAFE_CELL_CHARS)

    Dim i As Long
    For i = 1 To chunks.Count
        Dim target As Range
        Set target = startCell.Offset(i - 1, 0)

        If Len(CStr(chunks(i))) > MAX_CELL_CHARS Then
            Err.Raise vbObjectError + 1200, "WriteSpilledText", "A chunk exceeded Excel's 32,767 character limit."
        End If
        target.Value2 = chunks(i)
    Next i
End Sub

Public Function SplitIntoCellSizedChunks(ByVal sourceText As String, ByVal chunkSize As Long) As Collection
    Dim chunks As New Collection

    If Len(sourceText) = 0 Then
        chunks.Add vbNullString
        Set SplitIntoCellSizedChunks = chunks
        Exit Function
    End If

    Dim cursor As Long
    cursor = 1

    Do While cursor <= Len(sourceText)
        Dim remaining As Long
        remaining = Len(sourceText) - cursor + 1

        If remaining <= chunkSize Then
            chunks.Add Mid$(sourceText, cursor)
            Exit Do
        End If

        Dim window As String
        window = Mid$(sourceText, cursor, chunkSize)

        Dim splitPos As Long
        splitPos = FindBestSplitPosition(window)

        If splitPos <= 0 Then
            splitPos = chunkSize
        End If

        chunks.Add Left$(window, splitPos)
        cursor = cursor + splitPos

        Do While cursor <= Len(sourceText)
            Dim ch As String
            ch = Mid$(sourceText, cursor, 1)
            If ch = vbCr Or ch = vbLf Or ch = " " Or ch = "_" Then
                cursor = cursor + 1
            Else
                Exit Do
            End If
        Loop
    Loop

    Set SplitIntoCellSizedChunks = chunks
End Function

Private Function FindBestSplitPosition(ByVal textWindow As String) As Long
    Dim i As Long

    For i = Len(textWindow) To 1 Step -1
        If Mid$(textWindow, i, 1) = vbLf Then
            FindBestSplitPosition = i
            Exit Function
        End If
    Next i

    For i = Len(textWindow) To 1 Step -1
        If Mid$(textWindow, i, 1) = " " Then
            FindBestSplitPosition = i
            Exit Function
        End If
    Next i

    For i = Len(textWindow) To 1 Step -1
        If Mid$(textWindow, i, 1) = "_" Then
            FindBestSplitPosition = i
            Exit Function
        End If
    Next i

    FindBestSplitPosition = 0
End Function

Private Function HttpGet(ByVal url As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 1300, "HttpGet", "HTTP " & CStr(http.Status) & " from " & url & ": " & http.responseText
    End If

    HttpGet = CStr(http.responseText)
End Function

Private Function HttpPostJson(ByVal url As String, ByVal jsonPayload As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonPayload

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 1301, "HttpPostJson", "HTTP " & CStr(http.Status) & " from " & url & ": " & http.responseText
    End If

    HttpPostJson = CStr(http.responseText)
End Function

Private Function JsonEscape(ByVal value As String) As String
    Dim s As String
    s = value
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\"")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonEscape = s
End Function

Private Function JsonUnescape(ByVal value As String) As String
    Dim s As String
    s = value
    s = Replace(s, "\\", ChrW$(&HFFFF))
    s = Replace(s, "\n", vbLf)
    s = Replace(s, "\r", vbCr)
    s = Replace(s, "\t", vbTab)
    s = Replace(s, "\""", """)
    s = Replace(s, ChrW$(&HFFFF), "\")
    JsonUnescape = s
End Function

Private Function RegexFindAll(ByVal text As String, ByVal pattern As String) As Object
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.MultiLine = True
    re.IgnoreCase = True
    re.pattern = pattern

    Set RegexFindAll = re.Execute(text)
End Function
