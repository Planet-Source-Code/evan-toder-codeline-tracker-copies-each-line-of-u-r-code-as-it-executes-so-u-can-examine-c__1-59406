VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type


Public Enum EnumFilFlags
  FleReadOnly = &H1
  FleOverWritePrompt = &H2
  FleHideReadOnly = &H4
  FleNoChangeDir = &H8
  FleShowHelp = &H10
  FleEnableHook = &H20
  FleEnableTemplate = &H40
  FleEnableTemplateHandle = &H80
  FleNoValidate = &H100
  FleAllowMultiSelect = &H200
  FleExtensionDifferent = &H400
  FlePathMustExist = &H800
  FleFileMustExist = &H1000
  FleCreatePrompt = &H2000
  FleShareAware = &H4000
  FleNoReadOnlyReturn = &H8000
  FleNoTestFileCreate = &H10000
  FleNoNetworkButton = &H20000
  FleExplorer = &H80000
  FleLongnames = &H200000
End Enum
 
Private m_strDefaultExt As String
Private m_strDialogTitle As String
Private m_strFileName As String
Private m_strFileTitle As String
Private m_strInitialDir As String
Private m_strFilter As String
Private m_intFilterIndex As Integer
Private m_eFlags As EnumFilFlags
Private m_intMaxFileSize As Integer
Private m_lnghWndParent As Long

Private Const cintMaxFileLength As Integer = 260
 
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

'DEFAULT EXTENSION
Public Property Get DefaultExt() As String
  DefaultExt = m_strDefaultExt
End Property
Public Property Let DefaultExt(ByVal strValue As String)
  m_strDefaultExt = strValue
End Property

'DIALOG TITLE
Public Property Get DialogTitle() As String
  DialogTitle = m_strDialogTitle
End Property
Public Property Let DialogTitle(ByVal strValue As String)
  m_strDialogTitle = strValue
End Property

'FILENAME
Public Property Get FileName() As String
  FileName = m_strFileName
End Property
Public Property Let FileName(ByVal strValue As String)
  m_strFileName = strValue
End Property

'FILETITLE
Public Property Get FileTitle() As String
  FileTitle = m_strFileTitle
End Property
Public Property Let FileTitle(ByVal strValue As String)
  m_strFileTitle = strValue
End Property

'FILTER
Public Property Get Filter() As String
  Filter = m_strFilter
End Property
Public Property Let Filter(ByVal strValue As String)
  m_strFilter = strValue
End Property

'FILTER INDEX
Public Property Get FilterIndex() As Integer
  FilterIndex = m_intFilterIndex
End Property
Public Property Let FilterIndex(ByVal intValue As Integer)
  m_intFilterIndex = intValue
End Property

'FLAGS
Public Property Get Flags() As EnumFilFlags
  Flags = m_eFlags
End Property
Public Property Let Flags(ByVal eValue As EnumFilFlags)
  m_eFlags = eValue
End Property

'PARENT HWND
Public Property Get hWndParent() As Long
  hWndParent = m_lnghWndParent
End Property

Public Property Let hWndParent(ByVal lngValue As Long)
  m_lnghWndParent = lngValue
End Property

'INITIAL DIRECTORY
Public Property Get InitialDir() As String
  InitialDir = m_strInitialDir
End Property
Public Property Let InitialDir(ByVal strValue As String)
  m_strInitialDir = strValue
End Property

'MAX FILESIZE
Public Property Get MaxFileSize() As Integer
  MaxFileSize = m_intMaxFileSize
End Property
Public Property Let MaxFileSize(ByVal intValue As Integer)
  m_intMaxFileSize = intValue
End Property

Public Function FileShowSave(fOpen As Boolean) As Boolean
  On Error GoTo ERR_HANDLER:
'--------VARIABLES---------
  Dim of As OPENFILENAME
  Dim strChar As String * 1
  Dim intCounter As Integer
  Dim strTemp As String
 
'-----START CODE-----------
  With of ' Initialize the OPENFILENAME type
     .lpstrTitle = m_strDialogTitle & ""
     .Flags = m_eFlags
     .lpstrDefExt = m_strDefaultExt & ""
     .lStructSize = LenB(of)
     .lpstrFilter = m_strFilter & "||"
     .nFilterIndex = m_intFilterIndex
  
    ' To make Windows-style filter, replace pipes with nulls
     For intCounter = 1 To Len(m_strFilter)
        strChar = Mid$(m_strFilter, intCounter, 1)
        If strChar = "|" Then
          strTemp = strTemp & vbNullChar
        Else
          strTemp = strTemp & strChar
        End If
     Next
  
    ' Put double null at end
     strTemp = strTemp & vbNullChar & vbNullChar
     .lpstrFilter = strTemp
  
     ' Pad file and file title buffers to maximum path length
      strTemp = m_strFileName & String$(cintMaxFileLength - Len(m_strFileName), 0)
      
     .lpstrFile = strTemp
     .nMaxFile = cintMaxFileLength
  
     strTemp = m_strFileTitle & String$(cintMaxFileLength - Len(m_strFileTitle), 0)
     
     .lpstrFileTitle = strTemp
     .lpstrInitialDir = m_strInitialDir
     .nMaxFileTitle = cintMaxFileLength
     .hwndOwner = m_lnghWndParent
  End With
  
  ' If fOpen is true, show the Open file dialog, otherwise show the Save dialog
  If fOpen Then
    If GetOpenFileName(of) Then
      FileShowSave = True
      ' Assign property variables to appropriate values
      m_strFileName = TrimNulls(of.lpstrFile)
      m_strFileTitle = TrimNulls(of.lpstrFileTitle)
    Else
      FileShowSave = False
    End If
  Else
    If GetSaveFileName(of) Then
      FileShowSave = True
      ' Assign property variables to appropriate values
      m_strFileName = TrimNulls(of.lpstrFile)
      m_strFileTitle = TrimNulls(of.lpstrFileTitle)
    Else
      FileShowSave = False
    End If
  End If
Exit Function
'-----END CODE-------------
ERR_HANDLER:
  Debug.Print "FileShow  " & Err.Number; Err.Description
End Function

Private Function TrimNulls(ByVal strIn As String) As String
  On Error GoTo ERR_HANDLER:
  Dim intPos As Integer
'-----VARIABLES--------
  intPos = InStr(strIn, vbNullChar)
'-----BEGIN CODE-------
  If intPos = 0 Then
    ' No nulls in the string, just return it as is
    TrimNulls = strIn
  Else
    If intPos = 1 Then
      ' If the null character is at the first position, the
      ' entire string is a null string, so return a zero-length string
      TrimNulls = ""
    Else
      ' Not at the first position, so return the contents up
      ' to the occurrence of the null character
      TrimNulls = Left$(strIn, intPos - 1)
    End If
  End If
'-----END CODE----------
Exit Function
ERR_HANDLER:
  Debug.Print "ProcedureName  " & Err.Number; Err.Description
End Function


