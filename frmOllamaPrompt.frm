VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOllamaPrompt
   Caption         =   "Ollama Prompt Builder"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7365
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblInput
      Caption         =   "Input ranges (comma-separated):"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   3375
   End
   Begin VB.TextBox txtInputRanges
      Height          =   735
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   6975
   End
   Begin VB.Label lblOutput
      Caption         =   "Output start cell:"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   1395
      Width           =   1755
   End
   Begin VB.TextBox txtOutputCell
      Height          =   315
      Left            =   1935
      TabIndex        =   3
      Top             =   1365
      Width           =   1215
   End
   Begin VB.Label lblModel
      Caption         =   "Model:"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1845
      Width           =   855
   End
   Begin VB.ComboBox cboModel
      Height          =   315
      Left            =   1935
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1815
      Width           =   3375
   End
   Begin VB.CommandButton btnRefreshModels
      Caption         =   "Refresh"
      Height          =   345
      Left            =   5490
      TabIndex        =   6
      Top             =   1800
      Width           =   1665
   End
   Begin VB.Label lblStatus
      Caption         =   "Ready."
      Height          =   495
      Left            =   180
      TabIndex        =   7
      Top             =   2340
      Width           =   6975
      WordWrap        =   -1  'True
   End
   Begin VB.CommandButton btnRun
      Caption         =   "Run"
      Default         =   -1  'True
      Height          =   435
      Left            =   5265
      TabIndex        =   8
      Top             =   5655
      Width           =   915
   End
   Begin VB.CommandButton btnCancel
      Caption         =   "Close"
      Cancel          =   -1  'True
      Height          =   435
      Left            =   6270
      TabIndex        =   9
      Top             =   5655
      Width           =   915
   End
End
Attribute VB_Name = "frmOllamaPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    txtInputRanges.Text = "A1,A2:A4"
    txtOutputCell.Text = "B1"
    lblStatus.Caption = "Ready."
    RefreshModelList
End Sub

Private Sub btnRefreshModels_Click()
    RefreshModelList
End Sub

Private Sub RefreshModelList()
    On Error GoTo HandleErr

    cboModel.Clear

    Dim models As Collection
    Set models = OllamaSpillModule.GetAvailableModels()

    Dim i As Long
    For i = 1 To models.Count
        cboModel.AddItem CStr(models(i))
    Next i

    If cboModel.ListCount > 0 Then
        cboModel.ListIndex = 0
    End If

    lblStatus.Caption = "Loaded " & CStr(cboModel.ListCount) & " model(s)."
    Exit Sub

HandleErr:
    lblStatus.Caption = "Model load failed: " & Err.Description
End Sub

Private Sub btnRun_Click()
    On Error GoTo HandleErr

    lblStatus.Caption = "Building prompt..."
    DoEvents

    Dim prompt As String
    prompt = OllamaSpillModule.GetPromptFromRanges(txtInputRanges.Text)

    Dim modelName As String
    modelName = Trim$(cboModel.Value)

    lblStatus.Caption = "Querying Ollama model '" & modelName & "'..."
    DoEvents

    Dim responseText As String
    responseText = OllamaSpillModule.GenerateFromOllama(modelName, prompt)

    lblStatus.Caption = "Writing output to sheet..."
    DoEvents

    OllamaSpillModule.WriteSpilledText Trim$(txtOutputCell.Text), responseText

    lblStatus.Caption = "Done. Output written starting at " & Trim$(txtOutputCell.Text) & "."
    Exit Sub

HandleErr:
    lblStatus.Caption = "Error: " & Err.Description
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
