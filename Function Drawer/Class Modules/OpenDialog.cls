VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OpenDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type FileParts
    PathName As String
    FileTitle As String
End Type

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpStrFilter As String
        lpStrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpStrFile As String
        nMaxFile As Long
        lpStrFileTitle As String
        nMaxFileTitle As Long
        lpStrInitialDir As String
        lpStrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpStrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
        pvReserved As Long  'new in Windows 2000 and later
        dwReserved As Long  'new in Windows 2000 and later
        flagsEx As Long  'new in Windows 2000 and later
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Const MaxInteger As Long = (2 ^ 15) - 1

Const OFN_READONLY = &H1
Const OFN_HIDEREADONLY = &H4
Const OFN_NOCHANGEDIR = &H8
Const OFN_SHOWHELP = &H10
Const OFN_ENABLEHOOK = &H20
Const OFN_ENABLETEMPLATE = &H40
Const OFN_ENABLETEMPLATEHANDLE = &H80
Const OFN_NOVALIDATE = &H100
Const OFN_ALLOWMULTISELECT = &H200
Const OFN_EXTENSIONDIFFERENT = &H400
Const OFN_PATHMUSTEXIST = &H800
Const OFN_FILEMUSTEXIST = &H1000
Const OFN_CREATEPROMPT = &H2000
Const OFN_EXPLORER = &H80000                         '  new look commdlg
Const OFN_NODEREFERENCELINKS = &H100000
Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
Const OFN_ENABLESIZING = &H800000
Const OFN_DONTADDTORECENT = &H2000000
Const OFN_FORCESHOWHIDDEN = &H10000000
Const OFN_EX_NOPLACESBAR = &H1

Dim ofnOpenFileName As OPENFILENAME

Dim ofnPathName As String

Dim ofnFileName As String
Dim ofnFileNames() As String

Dim ofnInitialFileTitle As String
Dim ofnInitialDirectory As String
Dim ofnMaxFileNameSize As Long
Dim ofnMaxFileTitleSize As Long
Dim ofnAllowMultiSelect As Boolean
Dim ofnPromptToCreateFile As Boolean
Dim ofnDontAddToRecent As Boolean
Dim ofnMustTypeDefferantExtension As Boolean
Dim ofnFileNotExistWarning As Boolean
Dim ofnPathNotExistWarning As Boolean
Dim ofnShowHiddenFiles As Boolean
Dim ofnHideReadOnlyCheckBox As Boolean
Dim ofnNoChangeInitSelectedDirectory As Boolean
Dim ofnDontGetShortcutsTarget As Boolean
Dim ofnAllowInvalidChars As Boolean
Dim ofnCheckReadOnlyCheckBox As Boolean
Dim ofnShowHelpButton As Boolean
Dim ofnNoPlacesBar As Boolean
Dim ofnEnableSizing As Boolean
Dim ofnOk As Boolean
Dim ofnCancel As Boolean

Friend Property Get hwndOwner() As Long
    hwndOwner = ofnOpenFileName.hwndOwner
End Property

Friend Property Let hwndOwner(ByVal vNewValue As Long)
   ofnOpenFileName.hwndOwner = vNewValue
End Property

Friend Property Get InitialFileTitle() As String
    InitialFileTitle = ofnInitialFileTitle
End Property

Friend Property Let InitialFileTitle(ByVal vNewValue As String)
   ofnInitialFileTitle = vNewValue
End Property

Friend Property Get Filter() As String
    Filter = Replace(ofnOpenFileName.lpStrFilter, Chr$(0), "|")
End Property

Friend Property Let Filter(ByVal vNewValue As String)
   ofnOpenFileName.lpStrFilter = Replace(vNewValue, "|", Chr$(0))
End Property

Friend Property Get FilterIndex() As Long
    FilterIndex = ofnOpenFileName.nFilterIndex
End Property

Friend Property Let FilterIndex(ByVal vNewValue As Long)
   ofnOpenFileName.nFilterIndex = vNewValue
End Property

Friend Property Get FileName() As String
    FileName = Trim(Replace(ofnOpenFileName.lpStrFile, Chr$(0), ""))
End Property

Friend Property Get FileCount() As String
    FileCount = UBound(ofnFileNames) + 1
End Property

Public Function FileNames(nIndex As Integer) As String
    FileNames = ofnPathName & "\" & ofnFileNames(nIndex)
End Function

Public Function FileTitles(nIndex As Integer) As String
    FileTitles = ofnFileNames(nIndex)
    
End Function

Friend Property Get MaxFileNameSize() As Long
    MaxFileNameSize = ofnMaxFileNameSize
End Property

Friend Property Let MaxFileNameSize(ByVal vNewValue As Long)
   ofnMaxFileNameSize = vNewValue
End Property

Friend Property Get FileTitle() As String
    FileTitle = Trim(Replace(ofnOpenFileName.lpStrFileTitle, Chr$(0), ""))
End Property

Friend Property Get PathName() As String
    PathName = ofnPathName
End Property

Friend Property Get MaxFileTitleSize() As String
    MaxFileTitleSize = ofnMaxFileTitleSize
End Property

Friend Property Let MaxFileTitleSize(ByVal vNewValue As String)
   ofnMaxFileTitleSize = vNewValue
End Property

Friend Property Get InitialDirectory() As String
    InitialDirectory = ofnInitialDirectory
End Property

Friend Property Let InitialDirectory(ByVal vNewValue As String)
   ofnInitialDirectory = vNewValue
End Property

Friend Property Get DialogTitle() As String
    DialogTitle = ofnOpenFileName.lpStrTitle
End Property

Friend Property Let DialogTitle(ByVal vNewValue As String)
   ofnOpenFileName.lpStrTitle = vNewValue
End Property

Friend Property Get PromptToCreateFile() As Boolean
    PromptToCreateFile = ofnPromptToCreateFile
End Property

Friend Property Let PromptToCreateFile(ByVal vNewValue As Boolean)
   ofnPromptToCreateFile = vNewValue
End Property

Friend Property Get AllowMultiSelect() As Boolean
    AllowMultiSelect = ofnAllowMultiSelect
End Property

Friend Property Let AllowMultiSelect(ByVal vNewValue As Boolean)
   ofnAllowMultiSelect = vNewValue
End Property

Friend Property Get DontAddToRecent() As Boolean
    DontAddToRecent = ofnDontAddToRecent
End Property

Friend Property Let DontAddToRecent(ByVal vNewValue As Boolean)
   ofnDontAddToRecent = vNewValue
End Property

Friend Property Get MustTypeDefferantExtension() As Boolean
    MustTypeDefferantExtension = ofnMustTypeDefferantExtension
End Property

Friend Property Let MustTypeDefferantExtension(ByVal vNewValue As Boolean)
   ofnMustTypeDefferantExtension = vNewValue
End Property

Friend Property Get FileNotExistWarning() As Boolean
    FileNotExistWarning = ofnFileNotExistWarning
End Property

Friend Property Let FileNotExistWarning(ByVal vNewValue As Boolean)
   ofnFileNotExistWarning = vNewValue
End Property

Friend Property Get PathNotExistWarning() As Boolean
    PathNotExistWarning = ofnPathNotExistWarning
End Property

Friend Property Let PathNotExistWarning(ByVal vNewValue As Boolean)
   ofnPathNotExistWarning = vNewValue
End Property

Friend Property Get ShowHiddenFiles() As Boolean
    ShowHiddenFiles = ofnShowHiddenFiles
End Property

Friend Property Let ShowHiddenFiles(ByVal vNewValue As Boolean)
   ofnShowHiddenFiles = vNewValue
End Property

Friend Property Get HideReadOnlyCheckBox() As Boolean
    HideReadOnlyCheckBox = ofnHideReadOnlyCheckBox
End Property

Friend Property Let HideReadOnlyCheckBox(ByVal vNewValue As Boolean)
   ofnHideReadOnlyCheckBox = vNewValue
End Property
Friend Property Get NoChangeInitSelectedDirectory() As Boolean
    NoChangeInitSelectedDirectory = ofnNoChangeInitSelectedDirectory
End Property

Friend Property Let NoChangeInitSelectedDirectory(ByVal vNewValue As Boolean)
   ofnNoChangeInitSelectedDirectory = vNewValue
End Property

Friend Property Get DontGetShortcutsTarget() As Boolean
    DontGetShortcutsTarget = ofnDontGetShortcutsTarget
End Property

Friend Property Let DontGetShortcutsTarget(ByVal vNewValue As Boolean)
   ofnDontGetShortcutsTarget = vNewValue
End Property

Friend Property Get AllowInvalidChars() As Boolean
    AllowInvalidChars = ofnAllowInvalidChars
End Property

Friend Property Let AllowInvalidChars(ByVal vNewValue As Boolean)
   ofnAllowInvalidChars = vNewValue
End Property

Friend Property Get CheckReadOnlyCheckBox() As Boolean
    CheckReadOnlyCheckBox = ofnCheckReadOnlyCheckBox
End Property

Friend Property Let CheckReadOnlyCheckBox(ByVal vNewValue As Boolean)
   ofnCheckReadOnlyCheckBox = vNewValue
End Property

Friend Property Get ShowHelpButton() As Boolean
    ShowHelpButton = ofnShowHelpButton
End Property

Friend Property Let ShowHelpButton(ByVal vNewValue As Boolean)
   ofnShowHelpButton = vNewValue
End Property

Friend Property Get NoPlacesBar() As Boolean
    NoPlacesBar = ofnNoPlacesBar
End Property

Friend Property Let NoPlacesBar(ByVal vNewValue As Boolean)
   ofnNoPlacesBar = vNewValue
End Property

Friend Property Get EnableSizing() As Boolean
    EnableSizing = ofnEnableSizing
End Property

Friend Property Let EnableSizing(ByVal vNewValue As Boolean)
   ofnEnableSizing = vNewValue
End Property

Friend Property Get OK() As Boolean
    OK = ofnOk
End Property

Friend Property Get Cancel() As Boolean
    Cancel = ofnCancel
End Property

Public Sub ShowDialog()
    With ofnOpenFileName
        .Flags = Abs(ofnAllowMultiSelect * OFN_ALLOWMULTISELECT) Or _
                 Abs(ofnPromptToCreateFile * OFN_CREATEPROMPT) Or _
                 Abs(ofnDontAddToRecent * OFN_DONTADDTORECENT) Or _
                 Abs(ofnMustTypeDefferantExtension * OFN_EXTENSIONDIFFERENT) Or _
                 Abs(ofnFileNotExistWarning * OFN_FILEMUSTEXIST) Or _
                 Abs(ofnPathNotExistWarning * OFN_PATHMUSTEXIST) Or _
                 Abs(ofnShowHiddenFiles * OFN_FORCESHOWHIDDEN) Or _
                 Abs(ofnHideReadOnlyCheckBox * OFN_HIDEREADONLY) Or _
                 Abs(ofnCheckReadOnlyCheckBox * OFN_READONLY) Or _
                 Abs(ofnNoChangeInitSelectedDirectory * OFN_NOCHANGEDIR) Or _
                 Abs(ofnDontGetShortcutsTarget * OFN_NODEREFERENCELINKS) Or _
                 Abs(ofnAllowInvalidChars * OFN_NOVALIDATE) Or _
                 Abs(ofnShowHelpButton * OFN_SHOWHELP) Or _
                 Abs(ofnEnableSizing * OFN_ENABLESIZING) Or _
                 OFN_LONGNAMES Or OFN_EXPLORER
        .flagsEx = Abs(ofnNoPlacesBar * OFN_EX_NOPLACESBAR)
        .lStructSize = Len(ofnOpenFileName)
        .lpStrFile = ofnInitialFileTitle & Space$(ofnMaxFileNameSize) & vbNullChar
        .nMaxFile = Len(.lpStrFile)
        .lpStrFileTitle = vbNullChar & Space$(ofnMaxFileTitleSize) & vbNullChar & vbNullChar
        .nMaxFileTitle = Len(.lpStrFileTitle)
        .lpStrInitialDir = ofnInitialDirectory & vbNullChar & vbNullChar
    End With
    
    ofnOk = CBool(GetOpenFileName(ofnOpenFileName))
    ofnCancel = Not (ofnOk)
    
    ofnPathName = GetFileParts(Trim(Replace(ofnOpenFileName.lpStrFile, Chr$(0), ""))).PathName
    
    If ofnAllowMultiSelect = True Then
    
        Dim NewFileNames() As String
        Dim i As Integer
        
        
        ofnFileNames = Split(ofnOpenFileName.lpStrFile, Chr$(0))
        
        For i = 0 To UBound(ofnFileNames)
            ofnFileNames(i) = Trim(Replace(ofnFileNames(i), Chr$(0), ""))
            If ofnFileNames(i) <> "" Then
                ReDim Preserve NewFileNames(0 To i)
                NewFileNames(i) = ofnFileNames(i)
            End If
        Next
        
        If UBound(NewFileNames) > 0 Then
            ofnPathName = NewFileNames(0)
            
            For i = 0 To UBound(NewFileNames) - 1
                NewFileNames(i) = NewFileNames(i + 1)
            Next
            ReDim Preserve NewFileNames(UBound(NewFileNames) - 1)
        End If
        
        ofnFileNames = NewFileNames
    
    End If
    
End Sub
Private Sub Class_Initialize()
    ofnMaxFileNameSize = MaxInteger
    ofnMaxFileTitleSize = MaxInteger
    ofnEnableSizing = True
End Sub

Private Function GetFileParts(ByVal StrFileName As String) As FileParts
    On Error Resume Next
    Dim i As Integer
    Dim char As String
    For i = Len(StrFileName) To 1 Step -1
        char = Mid(StrFileName, i, 1)
        If char = "\" Or char = "/" Then Exit For
    Next
    GetFileParts.FileTitle = Right(StrFileName, Len(StrFileName) - i)
    GetFileParts.PathName = Left(StrFileName, i - 1)
End Function

