VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Code Hunter"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopyCode 
      Caption         =   "Copy"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7875
      TabIndex        =   7
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdStartSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   5940
      TabIndex        =   6
      Top             =   90
      Width           =   960
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   4815
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDisplayFunction 
      Height          =   1500
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3060
      Width           =   8700
   End
   Begin VB.CommandButton cmdCancelSearch 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6930
      TabIndex        =   3
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   375
      Left            =   5445
      TabIndex        =   2
      Top             =   90
      Width           =   420
   End
   Begin VB.TextBox txtFolderPath 
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   5280
   End
   Begin MSComctlLib.ListView lvResults 
      Height          =   2490
      Left            =   45
      TabIndex        =   0
      Top             =   495
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   4392
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "In Folder"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "In Function"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "At Line"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'api function to keep track of time
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public SearchMethod As VbCompareMethod
Public SearchString As String
Public FileMask     As String

Dim Search()        As String
Dim SearchCount     As Long

Private Sub cmdBrowse_Click()

  Dim strFolderPath As String
        
    'use txtFolderPath as our start path if we can
    If LenB(txtFolderPath.Text) > 2 Then
      'turn on error handling
      On Error Resume Next
      'check to see if the folder exists
      If Dir$(FormatPath(txtFolderPath.Text) & "\nul", vbDirectory) <> vbNullString Then
         strFolderPath = FormatPath(txtFolderPath.Text)
      Else
         strFolderPath = CurDir$
      End If
      'turn off error handling
      On Error GoTo 0
   End If
       
    'show the browse window
    strFolderPath = ShowBrowse(Me, "Select Code Folder", strFolderPath, False)
    
    'check to see if a folder was selected and fill in folder path textbox
    If Len(strFolderPath) > 0 Then
        txtFolderPath.Text = strFolderPath
        ChDir strFolderPath
    End If

End Sub

'the main search 'engine'
'gathers together all files and then looks through them
'each for a matching item for searchstring
Public Sub StartSearch()

  Dim lvListItem           As MSComctlLib.ListItem
  Dim colFiles             As Collection
  Dim strNextLine          As String
  Dim strCurrentFunction   As String
  Dim strCurrentCode       As String
  Dim lngCurrentLine       As Long
  Dim varCurrentFile       As Variant
  Dim intFilenum           As Integer
  Dim blnFoundItem         As Boolean
  Dim blnInsideFunction    As Boolean
  Dim lngTotalTime         As Long
  Dim lngTotalLines        As Long
  Dim lngTotalFiles        As Long
  Dim lngTotalBytes        As Double
    
    'check to make sure search string is valid
    If LenB(SearchString) = 0 Then
        MsgBox "Invalid Search String!", vbCritical
        GoTo Quit
    End If
    
    'make sure we have a file mask
    If LenB(FileMask) = 0 Then
        FileMask = "*.bas;*.frm;*.ctl"
    End If
    
    'clear list items if any
    lvResults.ListItems.Clear
    'clear function display text
    txtDisplayFunction.Text = ""
    
    'initialize and seperate search items into an array
    'to be used by the ContainsSearchItem() function
    If InStr(SearchString, "-") Then
        Search = Split(SearchString, "-")
        SearchCount = UBound(Search) + 1
      Else
        ReDim Search(0) As String
        Search(0) = SearchString
        SearchCount = 1
    End If
    
    'initialize variables that need it
    lngTotalTime = timeGetTime
    Set colFiles = New Collection
    
    'load all the files in the folder
    Status "Loading File List. Please Wait..."
    DoEvents
    Call RecurseFiles(colFiles, txtFolderPath, FileMask)
     
    'now loop through each file and search for matches to search items
    For Each varCurrentFile In colFiles
        Status CStr(varCurrentFile)
        
        lngTotalBytes = lngTotalBytes + FileLen(CStr(varCurrentFile))
        lngTotalFiles = lngTotalFiles + 1
        If lngTotalFiles Mod 5 = 0 Then DoEvents
        
        intFilenum = FreeFile
        lngCurrentLine = 0

        Open CStr(varCurrentFile) For Input As intFilenum
        Do While Not EOF(intFilenum)
            Line Input #intFilenum, strNextLine    'grab line of code
            
            lngTotalLines = lngTotalLines + 1
            lngCurrentLine = lngCurrentLine + 1
            
            If LenB(strNextLine) > 0 Then          'check length
                strNextLine = TabTrim(strNextLine) 'trim excess
                If LenB(strNextLine) > 0 Then      'check length again
                    
                    'quick check to see if it could be a function
                    If ContainsWord("Function ,Sub ,Property ", strNextLine) Then
                        'double check
                        If StartWord("Declare ,Private ,Public ,Friend ,Static ,Function ,Sub ,Property ", strNextLine) Then
                            'grab function name
                            strCurrentFunction = ParseFunctionName(strNextLine)
                            'store code
                            strCurrentCode = strNextLine & vbCrLf
                            'flip switch
                            blnInsideFunction = True
                            
                            'look for search items in current line of code
                            If ContainsSearchItem(strNextLine) Then
                                blnFoundItem = True
                            Else
                                blnFoundItem = False
                            End If
                            
                            'is it an API function?
                            If IsAPIFunction(strNextLine) Then
                                'check for line continuations
                                Do While RightCheck(strNextLine, "_") And Not EOF(intFilenum)
                                    'grab line of code
                                    Line Input #intFilenum, strNextLine
                                    'store it
                                    strCurrentCode = strCurrentCode & (strNextLine & vbCrLf)
                                    'increment line count
                                    lngTotalLines = lngTotalLines + 1
                                    If Cancelled Then GoTo Quit
                                Loop
                                'look for search items
                                blnFoundItem = ContainsSearchItem(strCurrentCode)
                            Else
                                'not an api call so its a regular function, sub, or property
                                'grab all the code for this code block
                                Do While blnInsideFunction And (Not EOF(intFilenum))
                                    'grab line of code
                                    Line Input #intFilenum, strNextLine
                                    'store it
                                    strCurrentCode = strCurrentCode & (strNextLine & vbCrLf)
                                    'increment line counter
                                    lngTotalLines = lngTotalLines + 1
                                    lngCurrentLine = lngCurrentLine + 1
                                    'check to see that its not a blank line
                                    If LenB(strNextLine) > 0 Then
                                        'look for end line
                                        If ContainsWord("Function,Sub,Property", strNextLine) Then
                                            If StartWord("End ", strNextLine) Then
                                                blnInsideFunction = False
                                            End If
                                        End If
                                    End If
                                    'check for search items in current line
                                    If ContainsSearchItem(strNextLine) Then
                                        blnFoundItem = True
                                    End If
                                    If Cancelled Then GoTo Quit
                                Loop
                            End If
                              
                            'did we find anything?
                            If blnFoundItem Then
                                'yes so add it to the list
                                blnFoundItem = False
                                Set lvListItem = lvResults.ListItems.Add(, , Mid$(CStr(varCurrentFile), InStrRev(CStr(varCurrentFile), "\") + 1))
                                With lvListItem
                                    'store the code in the tag property
                                    .Tag = Left$(strCurrentCode, Len(strCurrentCode) - 2)
                                    'file name
                                    .SubItems(1) = Left$(CStr(varCurrentFile), InStrRev(CStr(varCurrentFile), "\") - 1)
                                    'function name
                                    .SubItems(2) = strCurrentFunction
                                    'code line number
                                    .SubItems(3) = CStr(lngCurrentLine)
                                End With
                            End If
                        End If
                    End If
                End If
            End If
            If Cancelled Then GoTo Quit
        Loop

        lngCurrentLine = 0
        Close intFilenum

    Next varCurrentFile

Quit:
    'all done! :)
    cmdStartSearch.Enabled = True
    cmdCancelSearch.Enabled = False
    cmdCopyCode.Enabled = (lvResults.ListItems.Count > 0)
    Cancelled = False
    Status "Search Complete"
    
    'show some stats
    txtDisplayFunction.Text = "Files Searched" & vbTab & "= " & lngTotalFiles & vbCrLf & _
                              "Total Results" & vbTab & "= " & CStr(lvResults.ListItems.Count) & vbCrLf & _
                              "Total Bytes" & vbTab & "= " & FormatBytes(lngTotalBytes) & vbCrLf & _
                              "Total Lines" & vbTab & "= " & CStr(lngTotalLines) & vbCrLf & _
                              "Total Time" & vbTab & "= " & CStr(CalculateTime(timeGetTime - lngTotalTime)) & vbCrLf
    
End Sub

Private Sub cmdStartSearch_Click()

    frmFind.Show , Me

End Sub

Sub Status(ByVal St As String)

    StatusBar1.SimpleText = St

End Sub

Function ContainsSearchItem(ByVal strLine As String) As Boolean

  Dim X As Long

    For X = 0 To SearchCount - 1
        If InStr(1, strLine, Search(X), SearchMethod) Then
            ContainsSearchItem = True
            Exit Function
        End If
    Next X

End Function

Private Sub cmdCancelSearch_Click()

  Cancelled = True

End Sub

Private Sub cmdCopyCode_Click()
   
   If Not (lvResults.SelectedItem Is Nothing) Then
       Clipboard.Clear
       Clipboard.SetText lvResults.SelectedItem.Tag
   End If
   
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
   
   txtFolderPath = GetSetting(App.Title, "Settings", "LastFolder", "C:\")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  SaveSetting App.Title, "Settings", "LastFolder", txtFolderPath
  
  On Error Resume Next
  Unload frmFind
  
End Sub

Private Sub Form_Resize()

    On Error Resume Next

        If WindowState <> 1 Then
            With txtDisplayFunction
                .Height = Me.Height \ 4
                lvResults.Width = ScaleWidth - 100
                lvResults.Height = ScaleHeight - (550 + StatusBar1.Height + .Height)
                .Top = lvResults.Top + lvResults.Height + 50
                .Width = ScaleWidth - 100
                .Left = lvResults.Left
            End With
        End If

End Sub

Private Sub lvResults_DblClick()
    
       'open file if double clicked
       If Not (lvResults.SelectedItem Is Nothing) Then
           With lvResults.SelectedItem
               Shell "explorer.exe " & .SubItems(1) & "\" & .Text, vbNormalFocus
           End With
       End If
    
End Sub

Private Sub lvResults_ItemClick(ByVal Item As MSComctlLib.ListItem)

    txtDisplayFunction.Text = Item.Tag

    Status Search(0) & " Found in " & Item.Text & " in function " & Item.SubItems(2) & " @ line #" & Item.SubItems(3)

End Sub

