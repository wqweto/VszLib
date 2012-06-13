VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVszArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cVszArchive"

#Const Logging = VszLogging

'=========================================================================
' Implemented interfaces
'=========================================================================

Implements IVszArchiveOpenCallback
Implements IVszArchiveOpenVolumeCallback
'--- extract
Implements IVszArchiveExtractCallback
Implements IVszCryptoGetTextPassword
'--- compress
Implements IVszArchiveUpdateCallback
Implements IVszCryptoGetTextPassword2

'=========================================================================
' Public events
'=========================================================================

Event Progress(ByVal FileIdx As Long, ByVal Current As Double, ByVal Total As Double, Cancel As Boolean)
Event Error(Description As String, Source As String, Cancel As Boolean)
Event NewVolume(FileName As String)
Event BeforeExtract(ByVal FileIdx As Long, FileName As String, SkipFile As Boolean, Cancel As Boolean)

'=========================================================================
' API
'=========================================================================

'--- Variant types
Private Const VT_UI4                    As Long = 19
Private Const VT_UI8                    As Long = 21
Private Const VT_FILETIME               As Long = 64
'--- for FindFirstFile
Private Const MAX_PATH                  As Long = 260
Private Const INVALID_HANDLE_VALUE      As Long = -1
'--- for VirtualProtect
Private Const PAGE_EXECUTE_READWRITE    As Long = &H40
'--- for WaitForXxx
Private Const INFINITE                  As Long = -1
Private Const WAIT_OBJECT_0             As Long = 0
'--- registry
Private Const HKEY_LOCAL_MACHINE        As Long = &H80000002
Private Const SAM_READ                  As Long = &H20019
Private Const REG_SZ                    As Long = 1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As Any) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
'Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Private Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Private Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Private Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type SYSTEMTIME
    wYear               As Integer
    wMonth              As Integer
    wDayOfWeek          As Integer
    wDay                As Integer
    wHour               As Integer
    wMinute             As Integer
    wSecond             As Integer
    wMilliseconds       As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

Private Type CRITICAL_SECTION
    DebugInfo           As Long
    LockCount           As Long
    RecursionCount      As Long
    OwningThread        As Long
    LockSemaphore       As Long
    SpinCount           As Long
End Type
    
'=========================================================================
' Constants and member variables
'=========================================================================

Private Const ERR_LIBRARY       As String = "7-zip dll not found"
Private Const ERR_NOT_SUPPORTED As String = "Archive not supported"
Private Const ERR_CANCELLED     As String = "User cancelled"
Private Const ERR_OPEN_STREAM   As String = "Cannot open "
Private Const ERR_CREATE_STREAM As String = "Cannot create "
Private Const ERR_VOLUME_PROPERTY As String = "Volume property %1 not supported"
Private Const ERR_VOLUME_NOT_FOUND As String = "Volume %1 not found"

Private m_sLastError        As String
Private m_bCancel           As Boolean
Private m_sArchiveFile      As String
Private m_sPassword         As String
Private m_lCount            As Long
Private m_uFiles()          As VszFileInfo
Private m_lCurrentIdx       As Long
Private m_dblCurrent        As Double
Private m_dblTotal          As Double
Private m_oInArchive        As IVszInArchive
Private m_oInStream         As cVszStream
Private m_sTargetFolder     As String
Private m_oOutArchive       As IVszOutArchive
Private WithEvents m_oOutStream As cVszStream
Attribute m_oOutStream.VB_VarHelpID = -1
Private m_lThreadID         As Long
Private m_uThreadCallThunk  As VszCallThunk
Private m_aThreadParams()   As Long
Private m_uStdCallThunk     As VszCallThunk
Private m_hLibrary          As Long
Private m_uPfn              As VszLibraryExports
Private m_bSolid            As Boolean
Private m_oCacheStream      As cVszStream '--- used only in IDE
Private m_uThreadEvent      As VszThreadEventData
Private m_cCustomParams     As Collection
Private m_dblVolumeSize     As Double
Private m_bDirStructure     As Boolean

Private Type VszFileInfo
    FileName        As String
    Attributes      As Long
    Size            As Double
    Encrypted       As Boolean
    Crc             As Long
    Comment         As String
    CreationTime    As Date
    LastAccessTime  As Date
    LastWriteTime   As Date
    SourceFileName  As String
    Offset          As Double
End Type

Private Type VszCallThunk
    pfn             As Long
    Code(0 To 6)    As Long
End Type

Private Type VszLibraryExports
    CreateObject    As Long
'    GetNumberOfMethods As Long
'    GetMethodProperty As Long
    GetNumberOfFormats As Long
'    GetHandlerProperty As Long
    GetHandlerProperty2 As Long
'    SetLargePageMode As Long
End Type

Private Enum VszThreadEventTypeEnum
    vszEvtProgess = 1
    vszEvtError
    vszEvtNewVolume
    vszEvtBeforeExtract
End Enum

Private Type VszThreadEventData
    hLock           As CRITICAL_SECTION
    hFireEvent      As Long
    hCompleteEvent  As Long
    EventType       As VszThreadEventTypeEnum
    FileIdx         As Long
    Current         As Double
    Total           As Double
    Description     As String
    Source          As String
    FileName        As String
    SkipFile        As Boolean
    Cancel          As Boolean
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Function PrintError(sFunc As String, Optional ByVal SkipRaise As Boolean) As Boolean
    If Not m_bCancel Then
        m_sLastError = Err.Description
        Debug.Print MODULE_NAME & "." & sFunc & ": " & m_sLastError
        #If Logging Then
            DebugOutput "Error: " & Error$, MODULE_NAME & "." & sFunc & "(" & Erl & ")"
        #End If
        pvFireError m_sLastError, MODULE_NAME & "." & sFunc & "(" & Erl & ")", m_bCancel
    End If
    If m_bCancel And Not SkipRaise Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    PrintError = Not m_bCancel
End Function

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As String
    LastError = m_sLastError
End Property

Property Get Password() As String
    Password = m_sPassword
End Property

Property Let Password(sValue As String)
    m_sPassword = sValue
End Property

Property Get VolumeSize() As Double
    VolumeSize = m_dblVolumeSize
End Property

Property Let VolumeSize(dblValue As Double)
    m_dblVolumeSize = dblValue
End Property

Property Get FileCount() As Long
    FileCount = m_lCount
End Property

Property Get FileInfo(ByVal FileIdx As Long) As Variant
    If FileIdx >= 0 And FileIdx < m_lCount Then
        With m_uFiles(FileIdx)
            FileInfo = Array(.FileName, .Attributes, .Size, .Encrypted, .Crc, .Comment, .CreationTime, .LastAccessTime, .LastWriteTime)
        End With
    End If
End Property

Property Get FormatCount() As Long
    If m_hLibrary = 0 Then
        If Not Init() Then
            Err.Raise vbObjectError + 1, , ERR_LIBRARY
            Exit Property
        End If
    End If
    FormatCount = VszGetNumberOfFormats()
End Property

Property Get FormatInfo(ByVal FormatIdx As Long) As Variant
    Dim sClassID        As String
    Dim lPtr            As Long
    
    If m_hLibrary = 0 Then
        If Not Init() Then
            Err.Raise vbObjectError + 1, , ERR_LIBRARY
            Exit Property
        End If
    End If
    If FormatIdx >= 0 And FormatIdx < VszGetNumberOfFormats() Then
        lPtr = VszStringFromGUID(ByVal StrPtr(VszGetHandlerProperty2(FormatIdx, VszArchivePropID.vszClassID)))
        sClassID = SysAllocString(lPtr)
        Call CoTaskMemFree(lPtr)
        FormatInfo = Array(VszGetHandlerProperty2(FormatIdx, VszArchivePropID.vszName), _
            sClassID, _
            VszGetHandlerProperty2(FormatIdx, VszArchivePropID.vszExtension), _
            VszGetHandlerProperty2(FormatIdx, VszArchivePropID.vszAddExtension), _
            VszGetHandlerProperty2(FormatIdx, VszArchivePropID.vszUpdate), _
            VszGetHandlerProperty2(FormatIdx, VszArchivePropID.vszKeepName), _
            StrConv(VszGetHandlerProperty2(FormatIdx, VszArchivePropID.vszStartSignature), vbUnicode), _
            StrConv(VszGetHandlerProperty2(FormatIdx, VszArchivePropID.vszFinishSignature), vbUnicode))
    End If
End Property

Public Property Get Parameter(ParamName As String) As Variant
    On Error Resume Next
    Parameter = m_cCustomParams.Item(ParamName)(1)
    On Error GoTo 0
End Property

Public Property Let Parameter(ParamName As String, Value As Variant)
    Const FUNC_NAME     As String = "Parameter [let]"
    
    On Error GoTo EH
    If m_cCustomParams Is Nothing Then
        Set m_cCustomParams = New Collection
    End If
    On Error Resume Next
    m_cCustomParams.Remove ParamName
    On Error GoTo EH
    If Not IsEmpty(Value) Then
        Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle, vbDecimal
            Value = CLng(Value)
        Case Else
            Value = CStr(Value)
        End Select
        m_cCustomParams.Add Array(ParamName, Value), ParamName
    End If
    Exit Property
EH:
    If PrintError(FUNC_NAME, True) Then
        Resume Next
    End If
End Property

Property Get PreserveDirectoryStructure() As Boolean
    PreserveDirectoryStructure = m_bDirStructure
End Property

Property Let PreserveDirectoryStructure(ByVal bValue As Boolean)
    m_bDirStructure = bValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(Optional DllFile As String) As Boolean
    Dim sPath           As String
    
    Class_Terminate
    If LenB(DllFile) <> 0 Then
        m_hLibrary = LoadLibrary(DllFile)
    Else
        '--- auto load: 1. 7z.dll from app path, 2. 7z.dll from setup location, 3. 7za.dll from app path
        m_hLibrary = LoadLibrary(App.Path & "\7z.dll")
        If m_hLibrary = 0 Then
            sPath = pvRegRead("SOFTWARE\7-Zip", "Path")
            If LenB(sPath) <> 0 Then
                m_hLibrary = LoadLibrary(sPath & "\7z.dll")
            End If
        End If
        If m_hLibrary = 0 Then
            m_hLibrary = LoadLibrary(App.Path & "\7za.dll")
        End If
    End If
    If m_hLibrary <> 0 Then
        m_uPfn.CreateObject = GetProcAddress(m_hLibrary, "CreateObject")
'        m_uPfn.GetNumberOfMethods = GetProcAddress(m_hLibrary, "GetNumberOfMethods")
'        m_uPfn.GetMethodProperty = GetProcAddress(m_hLibrary, "GetMethodProperty")
        m_uPfn.GetNumberOfFormats = GetProcAddress(m_hLibrary, "GetNumberOfFormats")
'        m_uPfn.GetHandlerProperty = GetProcAddress(m_hLibrary, "GetHandlerProperty")
        m_uPfn.GetHandlerProperty2 = GetProcAddress(m_hLibrary, "GetHandlerProperty2")
'        m_uPfn.SetLargePageMode = GetProcAddress(m_hLibrary, "SetLargePageMode")
        '--- success
        Init = True
    End If
End Function

Public Function OpenArchive(ArchiveFile As String) As Boolean
    Const FUNC_NAME     As String = "OpenArchive"
    Const MAX_CHECK_POSITION As Long = &H10000
    Dim lIdx            As Long

    If m_hLibrary = 0 Then
        If Not Init() Then
            Err.Raise vbObjectError + 1, , ERR_LIBRARY
            Exit Function
        End If
    End If
    On Error GoTo EH
    m_sLastError = vbNullString
    m_bCancel = False
    m_sArchiveFile = ArchiveFile
    Set m_oInArchive = VszCreateObject(pvGetClassID(ArchiveFile), VszGUIDFromString(IID_IInArchive))
    If m_oInArchive Is Nothing Then
        Err.Raise vbObjectError + 1, , ERR_NOT_SUPPORTED
        Exit Function
    End If
    Set m_oInStream = pvInitStream(ArchiveFile)
    On Error Resume Next
    m_oInArchive.Open m_oInStream, MAX_CHECK_POSITION, Me
    On Error GoTo EH
    '--- check if open succeeded
    If m_oInArchive.GetNumberOfArchiveProperties() <> 0 Or m_oInArchive.GetNumberOfItems() Then
        m_bSolid = m_oInArchive.GetArchiveProperty(vszSolid)
        m_lCount = m_oInArchive.GetNumberOfItems()
        ReDim m_uFiles(0 To m_lCount) '--- last one empty
        With m_oInArchive
            For lIdx = 0 To m_lCount - 1
                m_uFiles(lIdx).FileName = .GetProperty(lIdx, vszPath)
                m_uFiles(lIdx).Size = .GetProperty(lIdx, vszSize) '--- VT_UI8
                m_uFiles(lIdx).Attributes = .GetProperty(lIdx, vszAttributes) Or -.GetProperty(lIdx, vszIsFolder) * vbDirectory
                m_uFiles(lIdx).Encrypted = .GetProperty(lIdx, vszEncrypted)
                m_uFiles(lIdx).Crc = .GetProperty(lIdx, vszCRC)
                m_uFiles(lIdx).Comment = .GetProperty(lIdx, vszComment)
                m_uFiles(lIdx).CreationTime = pvToDateVariant(.GetProperty(lIdx, vszCreationTime))
                m_uFiles(lIdx).LastAccessTime = pvToDateVariant(.GetProperty(lIdx, vszLastAccessTime))
                m_uFiles(lIdx).LastWriteTime = pvToDateVariant(.GetProperty(lIdx, vszLastWriteTime))
            Next
        End With
        '--- success
        OpenArchive = True
    End If
    Exit Function
EH:
    If PrintError(FUNC_NAME, True) Then
        Resume Next
    End If
End Function

Public Function Extract( _
            TargetFolder As String, _
            Optional Filter As Variant) As Boolean
    Const FUNC_NAME     As String = "Extract"
    Const IDX_INARCHIVE_EXTRACT As Long = 7
    Dim lIdx            As Long
    Dim sFileName       As String
    Dim aIndexes()      As Long
    Dim lIndexCount     As Long
    Dim lOffset         As Long
    Dim lInThis         As Long
    Dim pfnInExtract    As Long
    Dim pExtractCallback As IVszArchiveExtractCallback

    On Error GoTo EH
    m_sLastError = vbNullString
    If LenB(TargetFolder) <> 0 Then
        m_sTargetFolder = TargetFolder
    Else
        m_sTargetFolder = CurDir$()
    End If
    m_lCurrentIdx = -1
    m_dblCurrent = 0
    m_dblTotal = 0
    pfnInExtract = pvGetMemberPfn(VarPtr(m_oInArchive), IDX_INARCHIVE_EXTRACT, lInThis)
    Set pExtractCallback = Me
    ReDim aIndexes(-1 To -1) As Long
    If m_lCount > 0 And Not IsMissing(Filter) Then
        ReDim aIndexes(0 To m_lCount - 1) As Long
        If IsArray(Filter) Then
            For lIdx = 0 To m_lCount - 1
                If Not Filter(lIdx) Then
                Else
                    aIndexes(lIndexCount) = lIdx
                    lIndexCount = lIndexCount + 1
                End If
            Next
        ElseIf LenB(Filter) <> 0 Then
            For lIdx = 0 To m_lCount - 1
                sFileName = m_uFiles(lIdx).FileName
                If InStr(Filter, "\") = 0 Then
                    sFileName = Mid$(sFileName, InStrRev(sFileName, "\") + 1)
                End If
                If LCase$(sFileName) Like LCase$(Filter) Then
                    aIndexes(lIndexCount) = lIdx
                    lIndexCount = lIndexCount + 1
                End If
            Next
        Else
            lIndexCount = -1
        End If
    Else
        lIndexCount = -1
    End If
    If lIndexCount >= 0 Then
        For lIdx = 0 To m_lCount - 1
            m_uFiles(lIdx).Offset = -1
        Next
        For lIdx = 0 To lIndexCount - 1
            lOffset = lOffset + m_uFiles(aIndexes(lIdx)).Size
            m_uFiles(aIndexes(lIdx)).Offset = lOffset
        Next
        '--- solid archive -> use consecutive indexes
        If m_bSolid And lIndexCount > 0 Then
            For lIdx = 0 To aIndexes(lIndexCount - 1)
                aIndexes(lIdx) = lIdx
            Next
            lIndexCount = lIdx
        End If
        If InIde Then
            m_oInArchive.Extract aIndexes(0), lIndexCount, 0, Me
            Extract = True
        Else
            Extract = pvWaitThread(pvThreadCall(pfnInExtract, lInThis, VarPtr(aIndexes(0)), lIndexCount, 0, ObjPtr(pExtractCallback)))
        End If
    ElseIf m_lCount > 0 Then
        For lIdx = 0 To m_lCount - 1
            lOffset = lOffset + m_uFiles(lIdx).Size
            m_uFiles(lIdx).Offset = lOffset
        Next
        If InIde Then
            m_oInArchive.Extract ByVal 0, -1, 0, Me
            Extract = True
        Else
            Extract = pvWaitThread(pvThreadCall(pfnInExtract, lInThis, 0, -1, 0, ObjPtr(pExtractCallback)))
        End If
    End If
    Exit Function
EH:
    If PrintError(FUNC_NAME, True) Then
        Resume Next
    End If
End Function

Public Function AddFile( _
            File As String, _
            Optional Name As String, _
            Optional Comment As String) As Boolean
    Const FUNC_NAME     As String = "AddFile"
    Dim hFind           As Long
    Dim uData           As WIN32_FIND_DATA
    Dim cSize           As Currency

    On Error GoTo EH
    If m_lCount = 0 Then
        ReDim m_uFiles(0 To 2)
    ElseIf m_lCount > UBound(m_uFiles) Then
        ReDim Preserve m_uFiles(0 To 2 * UBound(m_uFiles))
    End If
    hFind = FindFirstFile(File, uData)
    If hFind <> INVALID_HANDLE_VALUE Then
        Call FindClose(hFind)
        Call CopyMemory(cSize, uData.nFileSizeLow, 4)
        Call CopyMemory(ByVal VarPtr(cSize) + 4, uData.nFileSizeHigh, 4)
    End If
    With m_uFiles(m_lCount)
        If LenB(Name) <> 0 Then
            .FileName = Name
        Else
            .FileName = Mid$(File, InStrRev(File, "\") + 1)
        End If
        .Size = cSize * 10000@
        .Attributes = uData.dwFileAttributes
        .Comment = Comment
        .CreationTime = pvToDate(uData.ftCreationTime)
        .LastAccessTime = pvToDate(uData.ftLastAccessTime)
        .LastWriteTime = pvToDate(uData.ftLastWriteTime)
        .SourceFileName = File
        If m_lCount > 0 Then
            .Offset = m_uFiles(m_lCount - 1).Offset + m_uFiles(m_lCount - 1).Size
        End If
    End With
    m_lCount = m_lCount + 1
    Exit Function
EH:
    If PrintError(FUNC_NAME, True) Then
        Resume Next
    End If
End Function

Public Function CompressArchive(ArchiveFile As String) As Boolean
    Const FUNC_NAME     As String = "CompressArchive"
    Const IDX_OUTARCHIVE_UPDATEITEMS As Long = 3
    Dim lOutThis         As Long
    Dim pfnOutUpdateItems As Long
    Dim pUpdateCallback As IVszArchiveUpdateCallback
    Dim pOutStream      As IVszSequentialOutStream
    Dim pSetter         As IVszSetProperties
    Dim vElem           As Variant
    Dim aNames()        As String
    Dim aValues()       As Variant
    Dim lIdx            As Long

    If m_hLibrary = 0 Then
        If Not Init() Then
            Err.Raise vbObjectError + 1, , ERR_LIBRARY
            Exit Function
        End If
    End If
    On Error GoTo EH
    m_sLastError = vbNullString
    m_bCancel = False
    m_lCurrentIdx = -1
    m_dblCurrent = 0
    m_dblTotal = 0
    Set m_oOutArchive = VszCreateObject(pvGetClassID(ArchiveFile), VszGUIDFromString(IID_IOutArchive))
    If m_oOutArchive Is Nothing Then
        Err.Raise vbObjectError + 1, , ERR_NOT_SUPPORTED
        Exit Function
    End If
    '--- set properties
    If Not m_cCustomParams Is Nothing Then
        ReDim aNames(0 To m_cCustomParams.Count)
        ReDim aValues(0 To m_cCustomParams.Count)
        For Each vElem In m_cCustomParams
            aNames(lIdx) = vElem(0)
            aValues(lIdx) = vElem(1)
            If VarType(aValues(lIdx)) = vbLong Then
                Call CopyMemory(aValues(lIdx), VT_UI4, 2)
            End If
            lIdx = lIdx + 1
        Next
        Set pSetter = m_oOutArchive
        pSetter.SetProperties aNames(0), aValues(0), lIdx
    End If
    pfnOutUpdateItems = pvGetMemberPfn(VarPtr(m_oOutArchive), IDX_OUTARCHIVE_UPDATEITEMS, lOutThis)
    Set pUpdateCallback = Me
    Set m_oOutStream = pvInitStream(ArchiveFile, OutputStream:=True, VolumeSize:=m_dblVolumeSize)
    Set pOutStream = m_oOutStream
    If InIde Then
        Set m_oCacheStream = pvInitStream()
        m_oOutArchive.UpdateItems pOutStream, m_lCount, Me
    Else
        CompressArchive = pvWaitThread(pvThreadCall(pfnOutUpdateItems, lOutThis, ObjPtr(pOutStream), m_lCount, ObjPtr(pUpdateCallback)))
    End If
    Set m_oOutArchive = Nothing
    '--- success
    CompressArchive = True
    Exit Function
EH:
    If PrintError(FUNC_NAME, True) Then
        Resume Next
    End If
End Function

'= private ===============================================================

Private Function VszCreateObject(clsid As VszGUID, iid As VszGUID) As IUnknown
    Dim lPtr            As Long

    Call pvStdCall(m_uPfn.CreateObject, VarPtr(clsid), VarPtr(iid), VarPtr(lPtr))
    Call CopyMemory(VszCreateObject, lPtr, 4)
End Function

Private Function VszGetNumberOfFormats() As Long
    Call pvStdCall(m_uPfn.GetNumberOfFormats, VarPtr(VszGetNumberOfFormats))
End Function

Private Function VszGetHandlerProperty2(ByVal index As Long, ByVal propID As VszArchivePropID) As Variant
    Call pvStdCall(m_uPfn.GetHandlerProperty2, index, propID, VarPtr(VszGetHandlerProperty2))
End Function

Private Function pvGetClassID(sFile As String) As VszGUID
    Dim lIdx            As Long
    Dim vElem           As Variant
    Dim sClassID        As String

    '--- first check for multi-volume case
    For lIdx = 0 To VszGetNumberOfFormats() - 1
        For Each vElem In Split(LCase$(VszGetHandlerProperty2(lIdx, VszArchivePropID.vszExtension)))
            If LCase$(sFile) Like "*." & LCase$(vElem) & ".001" Then
                GoTo Found
            End If
        Next
    Next
    For lIdx = 0 To VszGetNumberOfFormats() - 1
        For Each vElem In Split(LCase$(VszGetHandlerProperty2(lIdx, VszArchivePropID.vszExtension)))
            If LCase$(sFile) Like "*." & LCase$(vElem) Then
                GoTo Found
            End If
        Next
    Next
    Exit Function
Found:
    sClassID = VszGetHandlerProperty2(lIdx, VszArchivePropID.vszClassID)
    Call CopyMemory(pvGetClassID, ByVal StrPtr(sClassID), 16)
End Function

Private Function pvInitStream( _
            Optional sFile As String, _
            Optional ByVal OutputStream As Boolean, _
            Optional ByVal FileAttributes As Long, _
            Optional ByVal DateModified As Date, _
            Optional VolumeSize As Double, _
            Optional ByVal FileSize As Double, _
            Optional RetVal As cVszStream) As cVszStream
    Set RetVal = New cVszStream
    If OutputStream Then
        If Not RetVal.frInit(sFile, OutputStream:=True, FileAttributes:=FileAttributes, DateModified:=DateModified, MultiStream:=(VolumeSize <> 0&), VolumeSize:=VolumeSize, FileSize:=FileSize) Then
            #If Logging Then
                DebugOutput "Error creating " & sFile, "pvInitStream"
            #End If
            Err.Raise vbObjectError + 1, , ERR_CREATE_STREAM & sFile
        End If
    Else
        If Not RetVal.frInit(sFile, MultiStream:=(Right$(sFile, 4) = ".001")) Then
            Err.Raise vbObjectError + 1, , ERR_OPEN_STREAM & sFile
        End If
    End If
    Set pvInitStream = RetVal
End Function

Private Function MkPath(sPath As String) As Boolean
    On Error Resume Next
    MkPath = (GetAttr(sPath) And vbDirectory) <> 0
    If Not MkPath Then
        MkDir sPath
        MkPath = (GetAttr(sPath) And vbDirectory) <> 0
        If Not MkPath And InStrRev(sPath, "\") <> 0 Then
            MkPath Left$(sPath, InStrRev(sPath, "\") - 1)
            MkDir sPath
            MkPath = (GetAttr(sPath) And vbDirectory) <> 0
        End If
    End If
    On Error GoTo 0
End Function

Private Function pvToDate(uFileTime As FILETIME) As Date
    Dim uLocalTime      As FILETIME
    Dim uSysTime        As SYSTEMTIME

    Call FileTimeToLocalFileTime(uFileTime, uLocalTime)
    Call FileTimeToSystemTime(uLocalTime, uSysTime)
    With uSysTime
        pvToDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Private Function pvToDateVariant(vProp As Variant) As Date
    Dim nVarType        As Integer
    Dim uLocalTime      As FILETIME
    Dim uSysTime        As SYSTEMTIME

    If Not IsEmpty(vProp) Then
        Call CopyMemory(nVarType, vProp, 2)
        If nVarType = VT_FILETIME Then
            Call FileTimeToLocalFileTime(ByVal VarPtr(vProp) + 8, uLocalTime)
            Call FileTimeToSystemTime(uLocalTime, uSysTime)
            With uSysTime
                pvToDateVariant = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
            End With
        Else
            On Error Resume Next
            pvToDateVariant = CDate(vProp)
            On Error GoTo 0
        End If
    End If
End Function

Private Function pvToFileTime(ByVal dDate As Date) As Variant
    Dim uSysTime        As SYSTEMTIME
    Dim uFileTime       As FILETIME

    With uSysTime
        .wYear = Year(dDate)
        .wMonth = Month(dDate)
        .wDay = Day(dDate)
        .wHour = Hour(dDate)
        .wMinute = Minute(dDate)
        .wSecond = Second(dDate)
    End With
    Call SystemTimeToFileTime(uSysTime, uFileTime)
    Call LocalFileTimeToFileTime(uFileTime, ByVal VarPtr(pvToFileTime) + 8)
    Call CopyMemory(pvToFileTime, VT_FILETIME, 2)
End Function

Private Function pvRegRead(sSubKey As String, sName As String) As String
    Dim hKey            As Long
    Dim lType           As Long
    Dim lSize           As Long
    Dim sValue          As String
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sSubKey, 0, SAM_READ, hKey) = 0 Then
        If RegQueryValueEx(hKey, sName, 0, lType, ByVal 0&, lSize) = 0 Then
            If lType = REG_SZ Then
                sValue = String$(lSize, 0)
                If RegQueryValueEx(hKey, sName, 0, lType, ByVal sValue, lSize) = 0 Then
                    pvRegRead = sValue
                End If
            End If
        End If
        Call RegCloseKey(hKey)
    End If
End Function

Private Sub pvInitVbRuntime()
    Const FUNC_NAME     As String = "pvInitVbRuntime"
    Const PROGID_DUMMY  As String = LIB_NAME & ".cVszDummy"
    Dim lIdx            As Long

    lIdx = GetModuleHandle("MSVBVM60.DLL")
    lIdx = GetProcAddress(lIdx, "__vbaSetSystemError")
    Call RtlMoveMemory(lIdx, ByVal lIdx + 9, 4)
    Call RtlMoveMemory(lIdx, ByVal lIdx, 4)
    If TlsGetValue(lIdx) = 0 Then
        Call CoInitialize(0)
        Call CoCreateInstance(CLSIDFromProgID(PROGID_DUMMY), Nothing, CLSCTX_INPROC_SERVER, VszGUIDFromString(IID_IUnknown), Nothing)
        #If Logging Then
            DebugOutput "exit", FUNC_NAME
        #End If
    End If
End Sub

Private Function pvStdCall(ByVal pfn As Long, ParamArray A() As Variant) As Long
    Dim aParams()       As Long
    Dim lIdx            As Long
    Dim lArgs           As Long
'void _stdcall c_thunk(int pfn, int args, int count, int dummy) {
'    55                   push        ebp
'    8B EC                mov         ebp,esp
'    8B 4D 10             mov         ecx,dword ptr [count]
'    E3 09                jecxz       _skip_params
'    8B 55 0C             mov         edx,dword ptr [args]
'_params_loop:
'    FF 74 8A FC          push        dword ptr [edx+ecx*4-4]
'    E2 FA                loop        _params_loop
'_skip_params:
'    FF 55 08             call        dword ptr [pfn]
'    8B E5                mov         esp,ebp
'    5D                   pop         ebp
'    C2 10 00             ret         10h
'    90                   nop
'    90                   nop
'}
    If m_uStdCallThunk.pfn = 0 Then
        m_uStdCallThunk.Code(0) = &H8BEC8B55
        m_uStdCallThunk.Code(1) = &H9E3104D
        m_uStdCallThunk.Code(2) = &HFF0C558B
        m_uStdCallThunk.Code(3) = &HE2FC8A74
        m_uStdCallThunk.Code(4) = &H855FFFA
        m_uStdCallThunk.Code(5) = &HC25DE58B
        m_uStdCallThunk.Code(6) = &H90900010
        Call VirtualProtect(m_uStdCallThunk, Len(m_uStdCallThunk), PAGE_EXECUTE_READWRITE, lIdx)
        m_uStdCallThunk.pfn = VarPtr(m_uStdCallThunk.Code(0))
    End If
    If UBound(A) >= 0 Then
        ReDim aParams(0 To UBound(A)) As Long
        For lIdx = 0 To UBound(A)
            aParams(lIdx) = A(lIdx)
        Next
        lArgs = VarPtr(aParams(0))
    End If
    pvStdCall = CallWindowProc(m_uStdCallThunk.pfn, pfn, lArgs, UBound(A) + 1, 0)
End Function

Private Function pvThreadCall(ByVal pfn As Long, ParamArray A() As Variant) As Long
    Dim lIdx            As Long
'void d_thunk(int args) {
'   55                   push        ebp
'   8B EC                mov         ebp,esp
'   8B 55 08             mov         edx,dword ptr [args]
'   8B 4A 04             mov         ecx,dword ptr [edx+4]
'   E3 06                jecxz       _skip_params
'_params_loop:
'   FF 74 8A 04          push        dword ptr [edx+ecx*4+4]
'   E2 FA                loop        _params_loop
'_skip_params:
'   FF 12                call        dword ptr [edx]
'   8B E5                mov         esp,ebp
'   5D                   pop         ebp
'   C2 04 00             ret         4
'   90                   nop
'   90                   nop
'   90                   nop
'}
    If m_uThreadCallThunk.pfn = 0 Then
        m_uThreadCallThunk.Code(0) = &H8BEC8B55
        m_uThreadCallThunk.Code(1) = &H4A8B0855
        m_uThreadCallThunk.Code(2) = &HFF06E304
        m_uThreadCallThunk.Code(3) = &HE2048A74
        m_uThreadCallThunk.Code(4) = &H8B12FFFA
        m_uThreadCallThunk.Code(5) = &H4C25DE5
        m_uThreadCallThunk.Code(6) = &H90909000
        Call VirtualProtect(m_uThreadCallThunk, Len(m_uThreadCallThunk), PAGE_EXECUTE_READWRITE, lIdx)
        m_uThreadCallThunk.pfn = VarPtr(m_uThreadCallThunk.Code(0))
    End If
    ReDim m_aThreadParams(0 To UBound(A) + 2) As Long
    m_aThreadParams(0) = pfn
    m_aThreadParams(1) = UBound(A) + 1
    For lIdx = 0 To UBound(A)
        m_aThreadParams(lIdx + 2) = A(lIdx)
    Next
    pvThreadCall = CreateThread(0, 0, m_uThreadCallThunk.pfn, m_aThreadParams(0), 0, 0)
End Function

Private Function pvWaitThread(ByVal hThread As Long) As Boolean
    Const FUNC_NAME     As String = "pvWaitThread"
    Dim aHandles(0 To 1) As Long
    Dim lResult         As Long
    Dim lExitCode       As Long
    
    m_uThreadEvent.hFireEvent = CreateEvent(0, 0, 0, vbNullString)
    m_uThreadEvent.hCompleteEvent = CreateEvent(0, 0, 0, vbNullString)
    Call InitializeCriticalSection(m_uThreadEvent.hLock)
    m_uThreadEvent.Cancel = False
    aHandles(0) = hThread
    aHandles(1) = m_uThreadEvent.hFireEvent
    Do
        lResult = WaitForMultipleObjects(2, aHandles(0), 0, INFINITE)
        #If Logging Then
            DebugOutput "after WaitForMultipleObjects lResult=" & lResult & ", m_uThreadEvent.EventType=" & m_uThreadEvent.EventType, FUNC_NAME
        #End If
        If lResult = WAIT_OBJECT_0 + 1 Then
            With m_uThreadEvent
                Select Case .EventType
                Case vszEvtProgess
                    RaiseEvent Progress(.FileIdx, .Current, .Total, .Cancel)
                    Call SetEvent(.hCompleteEvent)
                Case vszEvtError
                    RaiseEvent Error(.Description, .Source, .Cancel)
                    Call SetEvent(.hCompleteEvent)
                Case vszEvtNewVolume
                    RaiseEvent NewVolume(.FileName)
                    Call SetEvent(.hCompleteEvent)
                Case vszEvtBeforeExtract
                    RaiseEvent BeforeExtract(.FileIdx, .FileName, .SkipFile, .Cancel)
                    Call SetEvent(.hCompleteEvent)
                End Select
            End With
        Else
            If lResult = WAIT_OBJECT_0 Then
                Call GetExitCodeThread(hThread, lExitCode)
                #If Logging Then
                    DebugOutput "lExitCode=&H" & Hex(lExitCode), FUNC_NAME
                #End If
                '--- success
                pvWaitThread = (lExitCode >= 0)
            End If
            Exit Do
        End If
    Loop
    Call CloseHandle(m_uThreadEvent.hFireEvent)
    Call CloseHandle(m_uThreadEvent.hCompleteEvent)
    Call DeleteCriticalSection(m_uThreadEvent.hLock)
End Function

Private Function pvGetMemberPfn(ByVal lPtr As Long, ByVal lIdx As Long, Optional lThisPtr As Long) As Long
    Call CopyMemory(lThisPtr, ByVal lPtr, 4)
    Call CopyMemory(pvGetMemberPfn, ByVal lThisPtr, 4)
    Call CopyMemory(pvGetMemberPfn, ByVal pvGetMemberPfn + lIdx * 4, 4)
End Function

Private Function pvFireProgress(ByVal FileIdx As Long, ByVal Current As Double, ByVal Total As Double, Cancel As Boolean) As Boolean
    If m_lThreadID = GetCurrentThreadId() Then
        RaiseEvent Progress(FileIdx, Current, Total, Cancel)
    Else
        With m_uThreadEvent
            Call EnterCriticalSection(.hLock)
            .EventType = vszEvtProgess
            .FileIdx = FileIdx
            .Current = Current
            .Total = Total
            .Cancel = Cancel
            Call SetEvent(.hFireEvent)
            Call WaitForMultipleObjects(1, .hCompleteEvent, 0, INFINITE)
            Cancel = .Cancel
            Call LeaveCriticalSection(.hLock)
        End With
    End If
    #If Logging Then
        If Cancel Then
            DebugOutput "Cancel=" & Cancel, "pvFireProgress"
        End If
    #End If
    pvFireProgress = Cancel
End Function

Private Function pvFireError(Description As String, Source As String, Cancel As Boolean) As Boolean
    If m_lThreadID = GetCurrentThreadId() Then
        RaiseEvent Error(Description, Source, Cancel)
    Else
        With m_uThreadEvent
            Call EnterCriticalSection(.hLock)
            .EventType = vszEvtError
            .Description = Description
            .Source = Source
            .Cancel = Cancel
            Call SetEvent(.hFireEvent)
            Call WaitForMultipleObjects(1, .hCompleteEvent, 0, INFINITE)
            Cancel = .Cancel
            Call LeaveCriticalSection(.hLock)
        End With
    End If
    pvFireError = Cancel
End Function

Private Function pvFireBeforeExtract(ByVal FileIdx As Long, FileName As String, SkipFile As Boolean, Cancel As Boolean) As Boolean
    If m_lThreadID = GetCurrentThreadId() Then
        RaiseEvent BeforeExtract(FileIdx, FileName, SkipFile, Cancel)
    Else
        With m_uThreadEvent
            Call EnterCriticalSection(.hLock)
            .EventType = vszEvtBeforeExtract
            .FileIdx = FileIdx
            .FileName = FileName
            .SkipFile = SkipFile
            .Cancel = Cancel
            Call SetEvent(.hFireEvent)
            Call WaitForMultipleObjects(1, .hCompleteEvent, 0, INFINITE)
            FileName = .FileName
            SkipFile = .SkipFile
            Cancel = .Cancel
            Call LeaveCriticalSection(.hLock)
        End With
    End If
    pvFireBeforeExtract = Cancel
End Function

Private Sub m_oOutStream_NotifyNewStream(FileName As String)
    If m_lThreadID = GetCurrentThreadId() Then
        RaiseEvent NewVolume(FileName)
    Else
        With m_uThreadEvent
            Call EnterCriticalSection(.hLock)
            .EventType = vszEvtNewVolume
            .FileName = FileName
            Call SetEvent(.hFireEvent)
            Call WaitForMultipleObjects(1, .hCompleteEvent, 0, INFINITE)
            Call LeaveCriticalSection(.hLock)
        End With
    End If
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_lThreadID = GetCurrentThreadId()
    m_bDirStructure = True
End Sub

Private Sub Class_Terminate()
    If Not m_oInArchive Is Nothing Or Not m_oOutArchive Is Nothing Then
        Set m_oInArchive = Nothing
        Set m_oOutArchive = Nothing
        '--- wait for worker threads to finish before unloading library DLL
        Call Sleep(100)
    End If
    If m_hLibrary <> 0 Then
        Call FreeLibrary(m_hLibrary)
        m_hLibrary = 0
    End If
End Sub

'=========================================================================
' IVszArchiveOpenCallback interface
'=========================================================================

Private Sub IVszArchiveOpenCallback_SetCompleted(ByVal files As VbSevenZip.INT_PTR, ByVal bytes As VbSevenZip.INT_PTR)
    #If Logging Then
        DebugOutput "exit", "IVszArchiveOpenCallback_SetCompleted"
    #End If
End Sub

Private Sub IVszArchiveOpenCallback_SetTotal(ByVal files As VbSevenZip.INT_PTR, ByVal bytes As VbSevenZip.INT_PTR)
    #If Logging Then
        DebugOutput "exit", "IVszArchiveOpenCallback_SetTotal"
    #End If
End Sub

'=========================================================================
' IVszArchiveOpenVolumeCallback interface
'=========================================================================

Private Function IVszArchiveOpenVolumeCallback_GetProperty(ByVal propID As VbSevenZip.VszItemPropID) As Variant
    Const FUNC_NAME     As String = "IVszArchiveOpenVolumeCallback_GetProperty"
    
    On Error GoTo EH
    #If Logging Then
        DebugOutput "propID=" & propID, "IVszArchiveOpenVolumeCallback_GetProperty"
    #End If
    Select Case propID
    Case VszItemPropID.vszName
        IVszArchiveOpenVolumeCallback_GetProperty = m_sArchiveFile
    Case VszItemPropID.vszSize
        If Not m_oInStream Is Nothing Then
            IVszArchiveOpenVolumeCallback_GetProperty = CCur(m_oInStream.frGetSize() / 10000@)
            Call CopyMemory(IVszArchiveOpenVolumeCallback_GetProperty, VT_UI8, 2)
        End If
    Case Else
        On Error GoTo 0
        Err.Raise vbObjectError + 1, , Replace(ERR_VOLUME_PROPERTY, "%1", propID)
    End Select
    Exit Function
EH:
    If PrintError(FUNC_NAME, True) Then
        Resume Next
    End If
End Function

Private Sub IVszArchiveOpenVolumeCallback_GetStream(ByVal name_ As VbSevenZip.INT_PTR, inStream As VbSevenZip.IVszInStream)
    Const FUNC_NAME     As String = "IVszArchiveOpenVolumeCallback_GetStream"
    Dim sFile           As String
    
    On Error GoTo EH
    If name_ <> 0 Then
        sFile = SysAllocString(name_)
    End If
    #If Logging Then
        DebugOutput "name=" & sFile, FUNC_NAME
    #End If
    '--- check for absolute path
    If Left$(sFile, 1) <> "\" And Mid$(sFile, 2, 1) <> ":" Then
        sFile = Left$(m_sArchiveFile, InStrRev(m_sArchiveFile, "\")) & sFile
    End If
    If Not FileExists(sFile) Then
        On Error GoTo 0
        Err.Raise vbObjectError + 1, , Replace(ERR_VOLUME_NOT_FOUND, "%1", sFile)
    End If
    Set inStream = pvInitStream(sFile)
    Set m_oInStream = inStream
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
    Exit Sub
EH:
    If PrintError(FUNC_NAME, True) Then
        Resume Next
    End If
End Sub

'=========================================================================
' IVszArchiveExtractCallback interface
'=========================================================================

Private Sub IVszArchiveExtractCallback_GetStream(ByVal index As Long, outStream As IVszSequentialOutStream, ByVal askExtractMode As VbSevenZip.VszAskMode)
    Const FUNC_NAME     As String = "IVszArchiveExtractCallback_GetStream"
    Dim sOutFile        As String
    Dim bSkipFile       As Boolean

    pvInitVbRuntime
    On Error GoTo EH
    #If Logging Then
        DebugOutput "index=" & index, FUNC_NAME
    #End If
    If m_uFiles(index).Offset >= 0 Then
        m_lCurrentIdx = index
        If LenB(m_uFiles(index).FileName) = 0 Then
            If m_lCount = 1 Then
                sOutFile = Mid$(m_sArchiveFile, InStrRev(m_sArchiveFile, "\") + 1)
                If InStrRev(sOutFile, ".") > 0 Then
                    sOutFile = Left$(sOutFile, InStrRev(sOutFile, ".") - 1)
                End If
                If LCase$(Right$(sOutFile, 4)) <> ".tar" Then
                    sOutFile = sOutFile & ".tar"
                End If
            Else
                sOutFile = "[no name] " & index
            End If
        Else
            sOutFile = m_uFiles(index).FileName
        End If
        If Not m_bDirStructure Then
            sOutFile = Mid$(sOutFile, InStrRev(sOutFile, "\") + 1)
        End If
        If Left$(sOutFile, 1) <> "\" And Mid$(sOutFile, 2, 1) <> ":" Then
            sOutFile = m_sTargetFolder & IIf(Right$(m_sTargetFolder, 1) <> "\", "\", vbNullString) & sOutFile
        End If
        '--- dont fire BeforeExtract for folders if not preserving directory structure
        If (m_uFiles(index).Attributes And vbDirectory) = 0 Or m_bDirStructure Then
            If pvFireBeforeExtract(index, sOutFile, bSkipFile, m_bCancel) Then
                #If Logging Then
                    DebugOutput "before ERR_CANCELLED, index=" & index, FUNC_NAME
                #End If
                On Error GoTo 0
                Err.Raise vbObjectError + 1, , ERR_CANCELLED
            End If
        End If
        If (m_uFiles(index).Attributes And vbDirectory) <> 0 Then
            If Not bSkipFile And m_bDirStructure Then
                MkPath sOutFile
            End If
        Else
            If Not bSkipFile Then
                MkPath Left$(sOutFile, InStrRev(sOutFile, "\"))
                Set outStream = pvInitStream(sOutFile, OutputStream:=True, FileAttributes:=m_uFiles(index).Attributes, DateModified:=m_uFiles(index).LastWriteTime, FileSize:=m_uFiles(index).Size)
            Else
                Set outStream = Nothing
            End If
        End If
    End If
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
    Exit Sub
EH:
    If PrintError(FUNC_NAME) Then
        Resume Next
    End If
End Sub

Private Sub IVszArchiveExtractCallback_PrepareOperation(ByVal askExtractMode As VbSevenZip.VszAskMode)
    pvInitVbRuntime
End Sub

Private Sub IVszArchiveExtractCallback_SetCompleted(ByVal completeValue As VbSevenZip.INT_PTR)
    Const FUNC_NAME     As String = "IVszArchiveExtractCallback_SetCompleted"
    Dim cValue          As Currency
    Dim lIdx            As Long

    pvInitVbRuntime
    #If Logging Then
        DebugOutput "enter", FUNC_NAME
    #End If
    If completeValue <> 0 Then
        Call CopyMemory(cValue, ByVal completeValue, 8)
        #If Logging Then
            DebugOutput "completeValue=" & cValue * 10000#, FUNC_NAME
        #End If
        m_dblCurrent = cValue * 10000#
        If pvFireProgress(-1, m_dblCurrent, m_dblTotal, m_bCancel) Then
            #If Logging Then
                DebugOutput "before ERR_CANCELLED ", FUNC_NAME
            #End If
            Err.Raise vbObjectError + 1, , ERR_CANCELLED
        End If
        For lIdx = 0 To m_lCount - 1
            If 0 <= m_uFiles(lIdx).Offset And m_uFiles(lIdx).Offset > m_dblCurrent Then
                If pvFireProgress(lIdx, m_dblCurrent - m_uFiles(lIdx).Offset + m_uFiles(lIdx).Size, m_uFiles(lIdx).Size, m_bCancel) Then
                    #If Logging Then
                        DebugOutput "before ERR_CANCELLED, lIdx=" & lIdx, FUNC_NAME
                    #End If
                    Err.Raise vbObjectError + 1, , ERR_CANCELLED
                End If
                Exit For
            End If
        Next
    End If
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

Private Sub IVszArchiveExtractCallback_SetOperationResult(ByVal result As VbSevenZip.VszExtOperationResult)
    Const FUNC_NAME     As String = "IVszArchiveExtractCallback_SetOperationResult"

    pvInitVbRuntime
    #If Logging Then
        DebugOutput "result=" & result, FUNC_NAME
    #End If
    Select Case result
    Case 0 '--- no error
        m_sLastError = vbNullString
    Case vszUnSupportedMethod
        m_sLastError = "Unsupported Method"
    Case vszDataError
        m_sLastError = "Data Error" & IIf(m_uFiles(m_lCurrentIdx).Encrypted, ". Wrong password?", vbNullString)
    Case vszCRCError
        m_sLastError = "CRC Error"
    Case Else
        m_sLastError = "Unknown error: " & result
    End Select
    If LenB(m_sLastError) <> 0 Then
        If Not m_bCancel Then
            pvFireError m_sLastError, MODULE_NAME & "." & FUNC_NAME, m_bCancel
        End If
        If m_bCancel Then
            Err.Raise Err.Number
        End If
    End If
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

Private Sub IVszArchiveExtractCallback_SetTotal(ByVal total_ As Currency)
    Const FUNC_NAME     As String = "IVszArchiveExtractCallback_SetTotal"

    pvInitVbRuntime
    #If Logging Then
        DebugOutput "total=" & total_ * 10000#, FUNC_NAME
    #End If
    m_dblTotal = total_ * 10000#
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

'=========================================================================
' IVszCryptoGetTextPassword interface
'=========================================================================

Private Function IVszCryptoGetTextPassword_CryptoGetTextPassword() As String
    Const FUNC_NAME     As String = "IVszCryptoGetTextPassword_CryptoGetTextPassword"

    pvInitVbRuntime
    #If Logging Then
        DebugOutput "m_sPassword=" & m_sPassword, FUNC_NAME
    #End If
    '--- bug in 7z.dll: cannot cast vbNullString to LPSTR
    IVszCryptoGetTextPassword_CryptoGetTextPassword = m_sPassword & ""
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Function

'=========================================================================
' IVszArchiveUpdateCallback interface
'=========================================================================

Private Function IVszArchiveUpdateCallback_GetProperty(ByVal index As Long, ByVal propID As VbSevenZip.VszItemPropID) As Variant
    Const FUNC_NAME     As String = "IVszArchiveUpdateCallback_GetProperty"

    pvInitVbRuntime
    #If Logging Then
        DebugOutput "index=" & index & ", propID=" & propID, FUNC_NAME
    #End If
    On Error GoTo EH
    Select Case propID
    Case vszIsAnti
        IVszArchiveUpdateCallback_GetProperty = False
    Case vszPath
        IVszArchiveUpdateCallback_GetProperty = m_uFiles(index).FileName
    Case VszItemPropID.vszExtension
        If InStrRev(m_uFiles(index).FileName, ".") > InStrRev(m_uFiles(index).FileName, "\") Then
            IVszArchiveUpdateCallback_GetProperty = Mid$(m_uFiles(index).FileName, InStrRev(m_uFiles(index).FileName, "\") + 1)
        End If
    Case vszSize
        IVszArchiveUpdateCallback_GetProperty = CCur(m_uFiles(index).Size / 10000@)
        Call CopyMemory(IVszArchiveUpdateCallback_GetProperty, VT_UI8, 2)
    Case vszAttributes
        IVszArchiveUpdateCallback_GetProperty = m_uFiles(index).Attributes
        Call CopyMemory(IVszArchiveUpdateCallback_GetProperty, VT_UI4, 2)
    Case vszIsFolder
        IVszArchiveUpdateCallback_GetProperty = (m_uFiles(index).Attributes And vbDirectory) <> 0
    Case vszEncrypted
        IVszArchiveUpdateCallback_GetProperty = LenB(m_sPassword) <> 0
    Case vszComment
        IVszArchiveUpdateCallback_GetProperty = m_uFiles(index).Comment
    Case vszCreationTime
        If m_uFiles(index).CreationTime <> 0 Then
            IVszArchiveUpdateCallback_GetProperty = pvToFileTime(m_uFiles(index).CreationTime)
        End If
    Case vszLastAccessTime
        If m_uFiles(index).LastAccessTime <> 0 Then
            IVszArchiveUpdateCallback_GetProperty = pvToFileTime(m_uFiles(index).LastAccessTime)
        End If
    Case vszLastWriteTime
        If m_uFiles(index).LastWriteTime <> 0 Then
            IVszArchiveUpdateCallback_GetProperty = pvToFileTime(m_uFiles(index).LastWriteTime)
        End If
    Case vszTimeType
        IVszArchiveUpdateCallback_GetProperty = VszFileTimeType.vszDOS
    Case Else
        #If Logging Then
            DebugOutput "unknown index=" & index & ", propID=" & propID, FUNC_NAME
        #End If
    End Select
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
    Exit Function
EH:
    If PrintError(FUNC_NAME, True) Then
        Resume Next
    End If
End Function

Private Sub IVszArchiveUpdateCallback_GetStream(ByVal index As Long, inStream As VbSevenZip.IVszSequentialInStream)
    Const FUNC_NAME     As String = "IVszArchiveUpdateCallback_GetStream"

    pvInitVbRuntime
    On Error GoTo EH
    #If Logging Then
        DebugOutput "index=" & index, FUNC_NAME
    #End If
    m_lCurrentIdx = index
    If m_oCacheStream Is Nothing Then
        #If Logging Then
            DebugOutput "before pvInitStream", FUNC_NAME
        #End If
        Set inStream = pvInitStream(m_uFiles(index).SourceFileName)
        If inStream Is Nothing Then
            Set inStream = pvInitStream() '--- null stream
        End If
    Else
        '--- note: possibly cross-thread call on Init method
        If Not m_oCacheStream.frInit(m_uFiles(index).SourceFileName) Then
            m_sLastError = ERR_OPEN_STREAM & m_uFiles(index).SourceFileName
            If pvFireError(m_sLastError, MODULE_NAME & "." & FUNC_NAME, m_bCancel) Then
                On Error GoTo 0
                Err.Raise vbObjectError + 1, , m_sLastError
            End If
            Exit Sub
        End If
        Set inStream = m_oCacheStream
    End If
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
    Exit Sub
EH:
    If PrintError(FUNC_NAME, True) Then
        Resume Next
    End If
End Sub

Private Sub IVszArchiveUpdateCallback_GetUpdateItemInfo(ByVal index As Long, newData As Long, newProperties As Long, indexInArchive As Long)
    Const FUNC_NAME     As String = "IVszArchiveUpdateCallback_GetUpdateItemInfo"
    
    pvInitVbRuntime
    #If Logging Then
        DebugOutput "enter", FUNC_NAME
    #End If
    newData = 1
    newProperties = 1
    indexInArchive = -1
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

Private Sub IVszArchiveUpdateCallback_SetCompleted(ByVal completeValue As VbSevenZip.INT_PTR)
    Const FUNC_NAME     As String = "IVszArchiveUpdateCallback_SetCompleted"
    Dim cValue          As Currency
    Dim lIdx            As Long

    pvInitVbRuntime
    #If Logging Then
        DebugOutput "enter", FUNC_NAME
    #End If
    If completeValue <> 0 Then
        Call CopyMemory(cValue, ByVal completeValue, 8)
        #If Logging Then
            DebugOutput "completeValue=" & cValue * 10000#, FUNC_NAME
        #End If
        m_dblCurrent = cValue * 10000#
        If pvFireProgress(-1, m_dblCurrent, m_dblTotal, m_bCancel) Then
            #If Logging Then
                DebugOutput "before ERR_CANCELLED", FUNC_NAME
            #End If
            Err.Raise vbObjectError + 1, , ERR_CANCELLED
        End If
        For lIdx = 0 To m_lCount - 1
            If m_uFiles(lIdx).Offset + m_uFiles(lIdx).Size >= m_dblCurrent Then
                If pvFireProgress(lIdx, m_dblCurrent - m_uFiles(lIdx).Offset, m_uFiles(lIdx).Size, m_bCancel) Then
                    #If Logging Then
                        DebugOutput "before ERR_CANCELLED, lIdx=" & lIdx, FUNC_NAME
                    #End If
                    Err.Raise vbObjectError + 1, , ERR_CANCELLED
                End If
                Exit For
            End If
        Next
    End If
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

Private Sub IVszArchiveUpdateCallback_SetOperationResult(ByVal operationResult As Long)
    Const FUNC_NAME     As String = "IVszArchiveUpdateCallback_SetOperationResult"
    
    pvInitVbRuntime
    #If Logging Then
        DebugOutput "operationResult=" & operationResult, FUNC_NAME
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

Private Sub IVszArchiveUpdateCallback_SetTotal(ByVal total_ As Currency)
    Const FUNC_NAME     As String = "IVszArchiveUpdateCallback_SetTotal"
    
    pvInitVbRuntime
    #If Logging Then
        DebugOutput "total=" & total_ * 10000@, FUNC_NAME
    #End If
    m_dblTotal = total_ * 10000#
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

'=========================================================================
' IVszCryptoGetTextPassword2 interface
'=========================================================================

Private Function IVszCryptoGetTextPassword2_CryptoGetTextPassword2(passwordIsDefined As Long) As String
    Const FUNC_NAME As String = "IVszCryptoGetTextPassword2_CryptoGetTextPassword2"
    
    pvInitVbRuntime
    #If Logging Then
        DebugOutput "m_sPassword=" & m_sPassword, FUNC_NAME
    #End If
    passwordIsDefined = -(LenB(m_sPassword) <> 0)
    '--- bug in 7z.dll: cannot cast vbNullString to LPSTR
    IVszCryptoGetTextPassword2_CryptoGetTextPassword2 = m_sPassword & ""
    #If Logging Then
        DebugOutput "passwordIsDefined=" & passwordIsDefined, FUNC_NAME
        DebugOutput "exit", FUNC_NAME
    #End If
End Function

