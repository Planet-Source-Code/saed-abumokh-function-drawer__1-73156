VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaveDialog"
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

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Const MaxInteger As Long = (2 ^ 15) - 1
Const OFN_OVERWRITEPROMPT = &H2
Const OFN_NOCHANGEDIR = &H8
Const OFN_SHOWHELP = &H10
Const OFN_ENABLEHOOK = &H20
Const OFN_ENABLETEMPLATE = &H40
Const OFN_ENABLETEMPLATEHANDLE = &H80
Const OFN_NOVALIDATE = &H100
Const OFN_EXTENSIONDIFFERENT = &H400
Const OFN_PATHMUSTEXIST = &H800
Const OFN_EXPLORER = &H80000                         '  new look commdlg
Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
Const OFN_ENABLESIZING = &H800000
Const OFN_DONTADDTORECENT = &H2000000
Const OFN_FORCESHOWHIDDEN = &H10000000

Const OFN_EX_NOPLACESBAR = &H1

Dim sfnSaveFileName As OPENFILENAME

Dim sfnPathName As String

Dim sfnFileName As String

Dim sfnInitialFileTitle As String
Dim sfnInitialDirectory As String
Dim sfnMaxFileNameSize As Long
Dim sfnMaxFileTitleSize As Long
Dim sfnDontAddToRecent As Boolean
Dim sfnMustTypeDefferantExtension As Boolean
Dim sfnPathNotExistWarning As Boolean
Dim sfnShowHiddenFiles As Boolean
Dim sfnNoChangeInitSelectedDirectory As Boolean
Dim sfnAllowInvalidChars As Boolean
Dim sfnShowHelpButton As Boolean
Dim sfnNoPlacesBar As Boolean
Dim sfnEnableSizing As Boolean
Dim sfnReplaceExistingFilePrompt As Boolean
Dim sfnAddExtension As Boolean
Dim sfnOk As Boolean
Dim sfnCancel As Boolean

Friend Property Get hwndOwner() As Long
    hwndOwner = sfnSaveFileName.hwndOwner
End Property

Friend Property Let hwndOwner(ByVal vNewValue As Long)
   sfnSaveFileName.hwndOwner = vNewValue
End Property

Friend Property Get InitialFileTitle() As String
    InitialFileTitle = sfnInitialFileTitle
End Property

Friend Property Let InitialFileTitle(ByVal vNewValue As String)
   sfnInitialFileTitle = vNewValue
End Property

Friend Property Get Filter() As String
    Filter = Replace(sfnSaveFileName.lpStrFilter, Chr$(0), "|")
End Property

Friend Property Let Filter(ByVal vNewValue As String)
   sfnSaveFileName.lpStrFilter = Replace(vNewValue, "|", Chr$(0))
End Property

Friend Property Get FilterIndex() As Long
    FilterIndex = sfnSaveFileName.nFilterIndex
End Property

Friend Property Let FilterIndex(ByVal vNewValue As Long)
   sfnSaveFileName.nFilterIndex = vNewValue
End Property

Friend Property Get FileName() As String
    FileName = sfnFileName
End Property

Friend Property Get MaxFileNameSize() As Long
    MaxFileNameSize = sfnMaxFileNameSize
End Property

Friend Property Let MaxFileNameSize(ByVal vNewValue As Long)
   sfnMaxFileNameSize = vNewValue
End Property

Friend Property Get FileTitle() As String
    FileTitle = Trim(Replace(sfnSaveFileName.lpStrFileTitle, Chr$(0), ""))
End Property

Friend Property Get PathName() As String
    PathName = sfnPathName
End Property

Friend Property Get MaxFileTitleSize() As String
    MaxFileTitleSize = sfnMaxFileTitleSize
End Property

Friend Property Let MaxFileTitleSize(ByVal vNewValue As String)
   sfnMaxFileTitleSize = vNewValue
End Property

Friend Property Get InitialDirectory() As String
    InitialDirectory = sfnInitialDirectory
End Property

Friend Property Let InitialDirectory(ByVal vNewValue As String)
   sfnInitialDirectory = vNewValue
End Property

Friend Property Get DialogTitle() As String
    DialogTitle = sfnSaveFileName.lpStrTitle
End Property

Friend Property Let DialogTitle(ByVal vNewValue As String)
   sfnSaveFileName.lpStrTitle = vNewValue
End Property

Friend Property Get DontAddToRecent() As Boolean
    DontAddToRecent = sfnDontAddToRecent
End Property

Friend Property Let DontAddToRecent(ByVal vNewValue As Boolean)
   sfnDontAddToRecent = vNewValue
End Property

Friend Property Get MustTypeDefferantExtension() As Boolean
    MustTypeDefferantExtension = sfnMustTypeDefferantExtension
End Property

Friend Property Let MustTypeDefferantExtension(ByVal vNewValue As Boolean)
   sfnMustTypeDefferantExtension = vNewValue
End Property

Friend Property Get PathNotExistWarning() As Boolean
    PathNotExistWarning = sfnPathNotExistWarning
End Property

Friend Property Let PathNotExistWarning(ByVal vNewValue As Boolean)
   sfnPathNotExistWarning = vNewValue
End Property

Friend Property Get ShowHiddenFiles() As Boolean
    ShowHiddenFiles = sfnShowHiddenFiles
End Property

Friend Property Let ShowHiddenFiles(ByVal vNewValue As Boolean)
   sfnShowHiddenFiles = vNewValue
End Property

Friend Property Get NoChangeInitSelectedDirectory() As Boolean
    NoChangeInitSelectedDirectory = sfnNoChangeInitSelectedDirectory
End Property

Friend Property Let NoChangeInitSelectedDirectory(ByVal vNewValue As Boolean)
   sfnNoChangeInitSelectedDirectory = vNewValue
End Property

Friend Property Get AllowInvalidChars() As Boolean
    AllowInvalidChars = sfnAllowInvalidChars
End Property

Friend Property Let AllowInvalidChars(ByVal vNewValue As Boolean)
   sfnAllowInvalidChars = vNewValue
End Property

Friend Property Get ReplaceExistingFilePrompt() As Boolean
    ReplaceExistingFilePrompt = sfnReplaceExistingFilePrompt
End Property

Friend Property Let ReplaceExistingFilePrompt(ByVal vNewValue As Boolean)
   sfnReplaceExistingFilePrompt = vNewValue
End Property

Friend Property Get AddExtension() As Boolean
    AddExtension = sfnAddExtension
End Property

Friend Property Let AddExtension(ByVal vNewValue As Boolean)
   sfnAddExtension = vNewValue
End Property

Friend Property Get ShowHelpButton() As Boolean
    ShowHelpButton = sfnShowHelpButton
End Property

Friend Property Let ShowHelpButton(ByVal vNewValue As Boolean)
   sfnShowHelpButton = vNewValue
End Property

Friend Property Get NoPlacesBar() As Boolean
    NoPlacesBar = sfnNoPlacesBar
End Property

Friend Property Let NoPlacesBar(ByVal vNewValue As Boolean)
   sfnNoPlacesBar = vNewValue
End Property

Friend Property Get EnableSizing() As Boolean
    EnableSizing = sfnEnableSizing
End Property

Friend Property Let EnableSizing(ByVal vNewValue As Boolean)
   sfnEnableSizing = vNewValue
End Property

Friend Property Get OK() As Boolean
    OK = sfnOk
End Property

Friend Property Get Cancel() As Boolean
    Cancel = sfnCancel
End Property

Public Sub ShowDialog()
    With sfnSaveFileName
        .Flags = Abs(sfnDontAddToRecent * OFN_DONTADDTORECENT) Or _
                 Abs(sfnMustTypeDefferantExtension * OFN_EXTENSIONDIFFERENT) Or _
                 Abs(sfnPathNotExistWarning * OFN_PATHMUSTEXIST) Or _
                 Abs(sfnReplaceExistingFilePrompt * OFN_OVERWRITEPROMPT) Or _
                 Abs(sfnShowHiddenFiles * OFN_FORCESHOWHIDDEN) Or _
                 Abs(sfnNoChangeInitSelectedDirectory * OFN_NOCHANGEDIR) Or _
                 Abs(sfnAllowInvalidChars * OFN_NOVALIDATE) Or _
                 Abs(sfnShowHelpButton * OFN_SHOWHELP) Or _
                 Abs(sfnEnableSizing * OFN_ENABLESIZING) Or _
                 OFN_LONGNAMES Or OFN_EXPLORER
        .flagsEx = Abs(sfnNoPlacesBar * OFN_EX_NOPLACESBAR)
        .lStructSize = Len(sfnSaveFileName)
        .lpStrFile = sfnInitialFileTitle & Space$(sfnMaxFileNameSize) & vbNullChar
        .nMaxFile = Len(.lpStrFile)
        .lpStrFileTitle = vbNullChar & Space$(sfnMaxFileTitleSize) & vbNullChar & vbNullChar
        .nMaxFileTitle = Len(.lpStrFileTitle)
        .lpStrInitialDir = sfnInitialDirectory & vbNullChar & vbNullChar
    End With
    
    sfnOk = CBool(GetSaveFileName(sfnSaveFileName))
    sfnCancel = Not (sfnOk)
    
    sfnFileName = (Trim(Replace(sfnSaveFileName.lpStrFile, Chr$(0), "")))
    If sfnAddExtension = True Then
        If Not (HasExtension(sfnFileName)) Then
            sfnFileName = sfnFileName & "." & GetExtensionFromFilter(sfnSaveFileName.lpStrFilter, sfnSaveFileName.nFilterIndex)
        End If
    End If
    
End Sub
Private Sub Class_Initialize()
    sfnMaxFileNameSize = MaxInteger
    sfnMaxFileTitleSize = MaxInteger
    sfnReplaceExistingFilePrompt = True
    sfnEnableSizing = True
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

Private Function HasExtension(ByVal StrFileName As String) As Boolean
    Dim str As String
    str = GetFileParts(StrFileName).FileTitle
    If str <> Replace(str, ".", "") Then HasExtension = True
End Function

Private Function GetExtensionFromFilter(ByVal Filter As String, ByVal FilterIndex As Long) As String
    Dim str() As String
    Dim Extension As String
    str = Split(Replace(Filter, Chr$(0), "|"), "|")
    Extension = str((FilterIndex * 2) - 1)
    Extension = Replace(Extension, "*.", "")
    Extension = Replace(Extension, "*", "")
    GetExtensionFromFilter = Extension
End Function
