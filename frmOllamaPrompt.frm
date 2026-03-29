VERSION 5.00
Begin VB.UserForm frmOllamaPrompt
   Caption         =   "Ollama Prompt Builder"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7365
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOllamaPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    BuildUi
    RefreshModelList
End Sub

Private Sub BuildUi()
    Dim topPos As Single
    topPos = 12

    AddLabel "lblInput", "Input ranges (comma-separated):", 12, topPos, 240, 18
    topPos = topPos + 18

    Dim txtIn As MSForms.TextBox
    Set txtIn = Me.Controls.Add("Forms.TextBox.1", "txtInputRanges")
    txtIn.Left = 12
    txtIn.Top = topPos
    txtIn.Width = 700
    txtIn.Height = 42
    txtIn.MultiLine = True
    txtIn.WordWrap = True
    txtIn.Text = "A1,A2:A4"

    topPos = topPos + 54
    AddLabel "lblOutput", "Output start cell:", 12, topPos, 120, 18

    Dim txtOut As MSForms.TextBox
    Set txtOut = Me.Controls.Add("Forms.TextBox.1", "txtOutputCell")
    txtOut.Left = 150
    txtOut.Top = topPos
    txtOut.Width = 120
    txtOut.Height = 20
    txtOut.Text = "B1"

    topPos = topPos + 30
    AddLabel "lblModel", "Model:", 12, topPos, 60, 18

    Dim cmb As MSForms.ComboBox
    Set cmb = Me.Controls.Add("Forms.ComboBox.1", "cboModel")
    cmb.Left = 150
    cmb.Top = topPos
    cmb.Width = 330
    cmb.Height = 20

    Dim btnRefresh As MSForms.CommandButton
    Set btnRefresh = Me.Controls.Add("Forms.CommandButton.1", "btnRefreshModels")
    btnRefresh.Left = 500
    btnRefresh.Top = topPos - 1
    btnRefresh.Width = 90
    btnRefresh.Height = 22
    btnRefresh.Caption = "Refresh"

    topPos = topPos + 30
    AddLabel "lblStatus", "Ready.", 12, topPos, 700, 30

    Dim btnRun As MSForms.CommandButton
    Set btnRun = Me.Controls.Add("Forms.CommandButton.1", "btnRun")
    btnRun.Left = 510
    btnRun.Top = 570
    btnRun.Width = 90
    btnRun.Height = 28
    btnRun.Caption = "Run"

    Dim btnCancel As MSForms.CommandButton
    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    btnCancel.Left = 615
    btnCancel.Top = 570
    btnCancel.Width = 90
    btnCancel.Height = 28
    btnCancel.Caption = "Close"
End Sub

Private Sub AddLabel(ByVal name As String, ByVal caption As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", name)
    lbl.Caption = caption
    lbl.Left = l
    lbl.Top = t
    lbl.Width = w
    lbl.Height = h
End Sub

Private Sub btnRefreshModels_Click()
    RefreshModelList
End Sub

Private Sub RefreshModelList()
    On Error GoTo HandleErr

    Dim cbo As MSForms.ComboBox
    Set cbo = Me.Controls("cboModel")
    cbo.Clear

    Dim models As Collection
    Set models = OllamaSpillModule.GetAvailableModels()

    Dim i As Long
    For i = 1 To models.Count
        cbo.AddItem CStr(models(i))
    Next i

    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If

    Me.Controls("lblStatus").Caption = "Loaded " & CStr(cbo.ListCount) & " model(s)."
    Exit Sub

HandleErr:
    Me.Controls("lblStatus").Caption = "Model load failed: " & Err.Description
End Sub

Private Sub btnRun_Click()
    On Error GoTo HandleErr

    Me.Controls("lblStatus").Caption = "Building prompt..."
    DoEvents

    Dim prompt As String
    prompt = OllamaSpillModule.GetPromptFromRanges(Me.Controls("txtInputRanges").Text)

    Dim modelName As String
    modelName = Trim$(Me.Controls("cboModel").Value)

    Me.Controls("lblStatus").Caption = "Querying Ollama model '" & modelName & "'..."
    DoEvents

    Dim responseText As String
    responseText = OllamaSpillModule.GenerateFromOllama(modelName, prompt)

    Me.Controls("lblStatus").Caption = "Writing output to sheet..."
    DoEvents

    OllamaSpillModule.WriteSpilledText Trim$(Me.Controls("txtOutputCell").Text), responseText

    Me.Controls("lblStatus").Caption = "Done. Output written starting at " & Trim$(Me.Controls("txtOutputCell").Text) & "."
    Exit Sub

HandleErr:
    Me.Controls("lblStatus").Caption = "Error: " & Err.Description
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
