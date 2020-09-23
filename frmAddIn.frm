VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My Add In"
   ClientHeight    =   3315
   ClientLeft      =   2175
   ClientTop       =   2235
   ClientWidth     =   5610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2205
      Top             =   1395
   End
   Begin CodelineTracker.button buttonStart 
      Height          =   510
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "start"
      Top             =   -45
      Width           =   530
      _extentx        =   926
      _extenty        =   900
      caption         =   ""
      font            =   "frmAddIn.frx":0000
      forecolor       =   0
      picture         =   "frmAddIn.frx":002C
      mouseovercolor  =   12648384
   End
   Begin CodelineTracker.ListGrid ListGrid1 
      Height          =   2940
      Left            =   90
      TabIndex        =   0
      Top             =   675
      Width           =   1755
      _extentx        =   3096
      _extenty        =   5186
      header_font     =   "frmAddIn.frx":0D06
      list_font       =   "frmAddIn.frx":0D34
      list_borderstyle=   0
   End
   Begin CodelineTracker.button buttonStop 
      Height          =   510
      Left            =   540
      TabIndex        =   2
      ToolTipText     =   "stop"
      Top             =   -45
      Width           =   530
      _extentx        =   926
      _extenty        =   900
      caption         =   ""
      font            =   "frmAddIn.frx":0D62
      forecolor       =   0
      picture         =   "frmAddIn.frx":0D8E
      showclickanimation=   -1  'True
      mouseovercolor  =   12632319
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsSub 
         Caption         =   "&minimize me"
         Index           =   0
      End
      Begin VB.Menu mnuOptionsSub 
         Caption         =   "&close me"
         Index           =   1
      End
   End
   Begin VB.Menu mnuSave 
      Caption         =   "&Save..."
      Begin VB.Menu mnuSaveSub 
         Caption         =   "s&ave to..."
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 
Public VBInstance As VBIDE.VBE
Public Connect As Connect
 
  

Private Sub buttonStart_Click()
    
    MsgBox "This project has not been coded to the point " & _
    "of being intelligent enouph to self-stop.  To stop " & _
    "copying code..press the red button.  Also, you can " & _
    "change the speed the code is copied via the timer"
    
    On Local Error GoTo local_error:
   ' disable this button
     buttonStart.Enabled = False
   ' put the focus to the code pane
1:   VBInstance.ActiveCodePane.Window.SetFocus
   ' timer does the work of copying code
2:   Timer1.Interval = 1000
     Timer1 = True
     
Exit Sub
local_error:
   ' if this error is raised it means the user attempted
   ' to start this program and didnt have a code window in view
   If Err.Number = 91 Then
     If Erl() = 1 Then
        MsgBox "Please change the active window to a code pane"
        Connect.Hide
        Exit Sub
     End If
   End If
End Sub

Private Sub buttonStop_Click()
   
   'stop the timer
   Timer1 = False
   ' break
   SendKeys "^ {BREAK}"
   ' enable the start button
   buttonStart.Enabled = True
   
End Sub

Private Sub Form_Load()
   
   ' set properties for the listgrid control
   With ListGrid1
       .Header = " code lines " & String(120, " ")
       .header_style = dretched
       .horizontal_scrollbar = True
       .list_borderstyle = lbsFlat
       .Move 0, 500
   End With
     
   ' set properties of start button
   With buttonStart
      .Move 0, 0, 525, 525
   End With
   
   ' set properties of end button
   With buttonStop
      .Move 530, 0, 525, 525
   End With
   
   'set some form propeties
   With Me
       .Width = ListGrid1.Width + 100
   End With
   
   
End Sub

Private Sub mnuOptionsSub_Click(Index As Integer)
  
Select Case Index
    Case Is = 0 ' minimize me
           Me.WindowState = vbMinimized
    Case Is = 1 ' clear listgrid and exit
           ListGrid1.clear
           Connect.Hide
    Case Is = 2

    Case Is = 3

End Select

End Sub

Private Sub mnuSaveSub_Click(Index As Integer)
   
 Select Case Index
    Case Is = 0 ' save contents of listgrid control (code)
        Dim classSave As New clsFile
        With classSave
           .Filter = "text|*.txt"
           .DialogTitle = "Save the code to..."
           .Flags = &H2 ' confirm overwrite
           .FileShowSave False
           ' if a valid filepath for saving contents
           ' of listgrid was specified then...
           If Len(.FileName) > 0 Then
              ' place contents of listgrid in variable
              Dim gridVals As String
              Dim l As Long
              ' row by row adding the next rowcontents
              For l = 0 To (ListGrid1.list_count - 1)
                 gridVals = (gridVals & ListGrid1.row_text(l) & vbCrLf)
              Next l
              ' now save contents of gridvals
              Dim ffile As Integer
              ffile = FreeFile
              Open .FileName For Output As #ffile
                 Print #ffile, gridVals
              Close #ffile
           End If
        End With
        
    Case Is = 1

    Case Is = 2

    Case Is = 3

End Select
  
   
End Sub

Private Sub Timer1_Timer()
     
    ' step to next line of code
    SendKeys "{F8}"
    'make sure were starting from beginning of line
    SendKeys "{HOME}", True
    ' hilight/select the current line of code
    SendKeys "+{END}", True
    ' copy it to clipboard
    SendKeys "^(c)", True
    ' place whats now in clipboard to variable
    Dim clipboardVal As String
    clipboardVal = Clipboard.GetText
    ' place it into the listbox
    ListGrid1.add_row clipboardVal
     
End Sub
