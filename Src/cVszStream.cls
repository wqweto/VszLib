VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVszStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cVszStream"

#Const Logging = VszLogging

'=========================================================================
' Implemented interfaces
'=========================================================================

Implements IVszSequentialInStream
Implements IVszInStream
Implements IVszSequentialOutStream
Implements IVszOutStream
Implements IVszOutStreamFlush
Implements IVszStreamGetSize

'=========================================================================
' Public events
'=========================================================================

Event NotifyNewStream(FileName As String)

'=========================================================================
' API
'=========================================================================

'--- for CreateFile
Private Const CREATE_ALWAYS         As Long = 2
Private Const OPEN_EXISTING         As Long = 3
Private Const OPEN_ALWAYS           As Long = 4
Private Const GENERIC_WRITE         As Long = &H40000000
Private Const GENERIC_READ          As Long = &H80000000
Private Const FILE_SHARE_READ       As Long = &H1
Private Const INVALID_HANDLE_VALUE  As Long = -1
'--- for SetFilePointer
Private Const FILE_BEGIN            As Long = 0
Private Const FILE_CURRENT          As Long = 1
Private Const FILE_END              As Long = 2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateFileA Lib "kernel32" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal NoSecurity As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal NoSecurity As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function SetFileAttributesA Lib "kernel32" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function SetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long

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

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const ERR_VOLUME_CREATE As String = "Cannot create volume %1"
Private Const ERR_REOPEN        As String = "Cannot reopen output volume %1"

Private m_hFile             As Long
Private m_cMultiStream      As Collection
Private m_dblVolumeSize     As Double
Private m_sFileName         As String
Private m_dblPosition       As Double
Private m_lFileAttributes   As Long
Private m_dDateModified     As Date

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String, Optional ByVal SkipRaise As Boolean)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Error$
    #If Logging Then
        DebugOutput "Error: " & Error$, MODULE_NAME & "." & sFunc & "(" & Erl & ")"
    #End If
    If Not SkipRaise Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

'=========================================================================
' Methods
'=========================================================================

Friend Function frInit( _
            sFileName As String, _
            Optional ByVal OutputStream As Boolean, _
            Optional ByVal MultiStream As Boolean, _
            Optional ByVal VolumeSize As Double, _
            Optional ByVal FileAttributes As Long, _
            Optional ByVal DateModified As Date, _
            Optional ByVal FileSize As Double) As Boolean
    Const FUNC_NAME     As String = "frInit"
    Dim oSubStream      As cVszStream
    Dim lIdx            As Long
    Dim dblOffset       As Double
    Dim sVolumeFile     As String

    On Error GoTo EH
    pvTerminate
    If MultiStream Then
        Set m_cMultiStream = New Collection
        If Right$(sFileName, 4) Like ".[0-9][0-9][0-9]" Then
            m_sFileName = Left$(sFileName, Len(sFileName) - 3)
        Else
            m_sFileName = sFileName & "."
        End If
        If OutputStream Then
            m_dblVolumeSize = VolumeSize
            pvAddNewStream
        Else
            For lIdx = 1 To 100000
                sVolumeFile = m_sFileName & Format$(lIdx, "000")
                If Not FileExists(sVolumeFile) Then
                    Exit For
                End If
                Set oSubStream = New cVszStream
                If Not oSubStream.frInit(sVolumeFile) Then
                    Exit For
                End If
                m_cMultiStream.Add Array(oSubStream, dblOffset, dblOffset + oSubStream.frGetSize(), sVolumeFile)
                dblOffset = dblOffset + oSubStream.frGetSize()
            Next
            '--- use base file as "single" multi-stream
            If lIdx = 1 Then
                Set oSubStream = New cVszStream
                If oSubStream.frInit(sFileName) Then
                    m_cMultiStream.Add Array(oSubStream, 0, oSubStream.frGetSize())
                End If
            End If
        End If
        frInit = (m_cMultiStream.Count > 0)
    Else
        m_sFileName = sFileName
        If OutputStream Then
            m_lFileAttributes = FileAttributes
            m_dDateModified = DateModified
            Call SetFileAttributes(sFileName, vbArchive)
            m_hFile = CreateFile(sFileName, GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_ALWAYS, vbArchive)
            If m_hFile <> INVALID_HANDLE_VALUE Then
                frSetSize FileSize / 10000@
            End If
        Else
            m_hFile = CreateFile(sFileName, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING)
        End If
        frInit = (m_hFile <> INVALID_HANDLE_VALUE) Or LenB(sFileName) = 0
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME, True
End Function

Friend Function frReopen(sFileName As String) As Boolean
    Const FUNC_NAME     As String = "frReopen"
    
    On Error GoTo EH
    If m_sFileName <> sFileName Then
        pvTerminate
        m_sFileName = sFileName
        Call SetFileAttributes(sFileName, vbArchive)
        m_hFile = CreateFile(sFileName, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_ALWAYS)
    End If
    frReopen = (m_hFile <> INVALID_HANDLE_VALUE)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Friend Function frGetSize() As Double
    Dim aQuad(0 To 1)   As Long
    Dim cSize           As Currency

    If m_cMultiStream Is Nothing Then
        aQuad(0) = GetFileSize(m_hFile, aQuad(1))
        Call CopyMemory(cSize, aQuad(0), 8)
        frGetSize = cSize * 10000#
    Else
        If m_cMultiStream.Count > 0 Then
            frGetSize = m_cMultiStream.Item(m_cMultiStream.Count)(2)
        End If
    End If
End Function

'= I/O ===================================================================

Friend Function frSeek(ByVal cDistance As Currency, ByVal lMoveMethod As Long, Optional ByVal lResultPtr As VbSevenZip.INT_PTR) As Currency
    Const FUNC_NAME     As String = "frRead"
    Dim aQuad(0 To 1)   As Long

    If m_cMultiStream Is Nothing Then
        Call CopyMemory(aQuad(0), cDistance, 8)
        aQuad(0) = SetFilePointer(m_hFile, aQuad(0), aQuad(1), lMoveMethod)
        Call CopyMemory(frSeek, aQuad(0), 8)
    Else
        If lMoveMethod = FILE_BEGIN Then
            m_dblPosition = 0
        ElseIf lMoveMethod = FILE_END Then
            m_dblPosition = frGetSize()
        End If
        m_dblPosition = m_dblPosition + cDistance * 10000#
        frSeek = m_dblPosition / 10000#
    End If
    If lResultPtr <> 0 Then
        Call CopyMemory(ByVal lResultPtr, frSeek, 8)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Friend Function frRead(ByVal lBuffer As Long, ByVal lSize As Long, Optional ByVal lResultPtr As VbSevenZip.INT_PTR) As Long
    Const FUNC_NAME     As String = "frRead"
    Dim dblOffset       As Double
    Dim lRead           As Long

    If m_cMultiStream Is Nothing Then
        If m_hFile <> INVALID_HANDLE_VALUE And lBuffer <> 0 Then
            Call ReadFile(m_hFile, ByVal lBuffer, lSize, frRead, 0)
        End If
        #If Logging Then
            DebugOutput "frRead=" & frRead & " " & m_sFileName, FUNC_NAME
        #End If
    Else
        Do
            With pvGetStream(m_dblPosition, dblOffset)
                If dblOffset < 0 Then
                    Exit Do
                End If
                .frSeek (m_dblPosition - dblOffset) / 10000@, FILE_BEGIN
                lRead = .frRead(lBuffer, lSize)
                m_dblPosition = m_dblPosition + lRead
                frRead = frRead + lRead
                lBuffer = lBuffer + lRead
                lSize = lSize - lRead
            End With
        Loop While lSize > 0 And lRead > 0
    End If
    If lResultPtr <> 0 Then
        Call CopyMemory(ByVal lResultPtr, frRead, 4)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME & "(m_dblPosition=" & m_dblPosition & ")"
    Resume Next
End Function

Friend Function frWrite(ByVal lBuffer As Long, ByVal lSize As Long, Optional ByVal lResultPtr As VbSevenZip.INT_PTR) As Long
    Const FUNC_NAME     As String = "frWrite"
    Dim dblOffset       As Double
    Dim lWritten        As Long

    On Error GoTo EH
    If m_cMultiStream Is Nothing Then
        If m_hFile <> INVALID_HANDLE_VALUE And lBuffer <> 0 Then
            Call WriteFile(m_hFile, ByVal lBuffer, lSize, frWrite, 0)
        Else
            frWrite = lSize
        End If
    Else
        Do
LoopAgain:
            With pvGetStream(m_dblPosition, dblOffset)
                If dblOffset < 0 Then
                    Exit Do
                End If
                .frSeek (m_dblPosition - dblOffset) / 10000@, FILE_BEGIN
                If m_dblVolumeSize > 0 And lSize >= m_dblVolumeSize - (m_dblPosition - dblOffset) Then
                    lWritten = .frWrite(lBuffer, m_dblVolumeSize - (m_dblPosition - dblOffset))
                    If Not pvAddNewStream() Then
                        On Error GoTo 0
                        Err.Raise vbObjectError + 2, , Replace(ERR_VOLUME_CREATE, "%1", m_cMultiStream.Count)
                    End If
                Else
                    lWritten = .frWrite(lBuffer, lSize)
                End If
                m_dblPosition = m_dblPosition + lWritten
                frWrite = frWrite + lWritten
                lBuffer = lBuffer + lWritten
                lSize = lSize - lWritten
            End With
        Loop While lSize > 0 And lWritten > 0
    End If
    If lResultPtr <> 0 Then
        Call CopyMemory(ByVal lResultPtr, frWrite, 4)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME & "(m_dblPosition=" & m_dblPosition & ")"
    Resume Next
End Function

Friend Sub frSetSize(ByVal cSize As Currency)
    Dim cCurrent        As Currency
    Dim dblOffset       As Double

    If m_cMultiStream Is Nothing Then
        cCurrent = frSeek(0, FILE_CURRENT)
        Call frSeek(cSize, FILE_BEGIN)
        Call SetEndOfFile(m_hFile)
        frSeek cCurrent, FILE_BEGIN
    Else
        With pvGetStream(cSize * 10000#, dblOffset)
            If dblOffset >= 0 Then
                .frSetSize cSize - dblOffset / 10000@
            End If
        End With
    End If
End Sub

'= private ===============================================================

Private Sub pvTerminate()
    If m_hFile <> INVALID_HANDLE_VALUE Then
        If m_dDateModified <> 0 Then
            Call SetFileTime(m_hFile, pvToFileTime(m_dDateModified), pvToFileTime(m_dDateModified), pvToFileTime(m_dDateModified))
        End If
        Call CloseHandle(m_hFile)
        m_hFile = INVALID_HANDLE_VALUE
        If m_lFileAttributes <> 0 And LenB(m_sFileName) <> 0 Then
            Call SetFileAttributes(m_sFileName, m_lFileAttributes)
        End If
    End If
    Set m_cMultiStream = Nothing
    m_dblVolumeSize = 0
    m_sFileName = vbNullString
    m_dblPosition = 0
    m_lFileAttributes = 0
    m_dDateModified = 0
End Sub

Private Function pvGetStream(dblPosition As Double, dblOffset As Double) As cVszStream
    Const FUNC_NAME     As String = "pvGetStream"
    Dim vElem           As Variant
    
    On Error GoTo EH
    For Each vElem In m_cMultiStream
        If dblPosition >= vElem(1) And dblPosition < vElem(2) Then
            Set pvGetStream = vElem(0)
            If Not pvGetStream.frReopen(CStr(vElem(3))) Then
                On Error GoTo 0
                Err.Raise vbObjectError + 2, , Replace(ERR_REOPEN, "%1", vElem(3))
            End If
            dblOffset = vElem(1)
            GoTo QH
        End If
    Next
    If pvGetStream Is Nothing Then
        dblOffset = -1
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvAddNewStream() As Boolean
    Const FUNC_NAME     As String = "pvAddNewStream"
    Dim dblOffset       As Double
    Dim sVolumeFile     As String
    Dim oSubStream      As cVszStream
    
    On Error GoTo EH
    dblOffset = frGetSize()
    sVolumeFile = m_sFileName & Format$(m_cMultiStream.Count + 1, "000")
    If m_cMultiStream.Count = 0 Then
        Set oSubStream = New cVszStream
        If Not oSubStream.frInit(sVolumeFile, OutputStream:=True) Then
            GoTo QH
        End If
    Else
        '--- hack: reuse stream on main worker thread
        Set oSubStream = m_cMultiStream.Item(1)(0)
        oSubStream.frReopen sVolumeFile
        oSubStream.frSetSize 0
    End If
    #If Logging Then
        DebugOutput "created " & sVolumeFile, FUNC_NAME
    #End If
    RaiseEvent NotifyNewStream(sVolumeFile)
    m_cMultiStream.Add Array(oSubStream, dblOffset, dblOffset + m_dblVolumeSize, sVolumeFile)
    '--- success
    pvAddNewStream = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvToFileTime(ByVal dCreated As Date) As FILETIME
    Dim uSysTime        As SYSTEMTIME
    Dim uFileTime       As FILETIME

    With uSysTime
        .wYear = Year(dCreated)
        .wMonth = Month(dCreated)
        .wDay = Day(dCreated)
        .wHour = Hour(dCreated)
        .wMinute = Minute(dCreated)
        .wSecond = Second(dCreated)
    End With
    Call SystemTimeToFileTime(uSysTime, uFileTime)
    Call LocalFileTimeToFileTime(uFileTime, pvToFileTime)
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

Private Function CreateFile(ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal NoSecurity As Long, ByVal dwCreationDisposition As Long, Optional ByVal dwFlagsAndAttributes As Long, Optional ByVal hTemplateFile As Long) As Long
    If IsNT Then
        CreateFile = CreateFileW(StrPtr(lpFileName), dwDesiredAccess, dwShareMode, NoSecurity, dwCreationDisposition, dwFlagsAndAttributes, hTemplateFile)
    Else
        CreateFile = CreateFileA(lpFileName, dwDesiredAccess, dwShareMode, NoSecurity, dwCreationDisposition, dwFlagsAndAttributes, hTemplateFile)
    End If
End Function

Private Function SetFileAttributes(ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
    If IsNT Then
        SetFileAttributes = SetFileAttributesW(StrPtr(lpFileName), dwFileAttributes)
    Else
        SetFileAttributes = SetFileAttributesA(lpFileName, dwFileAttributes)
    End If
End Function

Private Function IsNT() As Boolean
    Static lVersion     As Long
    
    If lVersion = 0 Then
        lVersion = GetVersion()
    End If
    IsNT = lVersion >= 0
End Function

'=========================================================================
' IVszSequentialInStream interface
'=========================================================================

Private Sub IVszSequentialInStream_Read(ByVal Data As VbSevenZip.INT_PTR, ByVal size_ As Long, ByVal processedSize As VbSevenZip.INT_PTR)
    Const FUNC_NAME     As String = "IVszSequentialInStream_Read "
    Dim lSize           As Long

    pvInitVbRuntime
    #If Logging Then
        DebugOutput "size=" & size_ & " " & m_sFileName, FUNC_NAME
    #End If
    frRead Data, size_, processedSize
    If processedSize <> 0 Then
        Call CopyMemory(lSize, ByVal processedSize, 4)
        #If Logging Then
            DebugOutput "processedSize=" & lSize, FUNC_NAME
        #End If
    End If
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

'=========================================================================
' IVszInStream interface
'=========================================================================

Private Sub IVszInStream_Read(ByVal Data As VbSevenZip.INT_PTR, ByVal size_ As Long, ByVal processedSize As VbSevenZip.INT_PTR)
    Const FUNC_NAME     As String = "IVszInStream_Read"
    Dim lSize           As Long

    pvInitVbRuntime
    #If Logging Then
        DebugOutput "size=" & size_ & " " & m_sFileName, FUNC_NAME
    #End If
    frRead Data, size_, processedSize
    If processedSize <> 0 Then
        Call CopyMemory(lSize, ByVal processedSize, 4)
        #If Logging Then
            DebugOutput "processedSize=" & lSize, FUNC_NAME
        #End If
    End If
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

Private Sub IVszInStream_Seek(ByVal offset_ As Currency, ByVal seekOrigin As Long, ByVal newPosition As VbSevenZip.INT_PTR)
    Const FUNC_NAME     As String = "IVszInStream_Seek"

    pvInitVbRuntime
    #If Logging Then
        DebugOutput "offset=" & offset_ * 10000@ & ", seekOrigin=" & seekOrigin & " " & m_sFileName, FUNC_NAME
    #End If
    #If Logging Then
        DebugOutput "retval=" & frSeek(offset_, seekOrigin, newPosition) * 10000@, FUNC_NAME
    #Else
        frSeek offset_, seekOrigin, newPosition
    #End If
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

'=========================================================================
' IVszSequentialOutStream interface
'=========================================================================

Private Sub IVszSequentialOutStream_Write(ByVal Data As VbSevenZip.INT_PTR, ByVal size_ As Long, ByVal processedSize As VbSevenZip.INT_PTR)
    Const FUNC_NAME     As String = "IVszSequentialOutStream_Write"

    pvInitVbRuntime
    #If Logging Then
        DebugOutput "size=" & size_, FUNC_NAME
    #End If
    frWrite Data, size_, processedSize
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

'=========================================================================
' IVszOutStream interface
'=========================================================================

Private Sub IVszOutStream_Write(ByVal Data As VbSevenZip.INT_PTR, ByVal size_ As Long, ByVal processedSize As VbSevenZip.INT_PTR)
    Const FUNC_NAME     As String = "IVszOutStream_Write"

    pvInitVbRuntime
    #If Logging Then
        DebugOutput "size=" & size_, FUNC_NAME
    #End If
    frWrite Data, size_, processedSize
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

Private Sub IVszOutStream_Seek(ByVal offset_ As Currency, ByVal seekOrigin As Long, ByVal newPosition As VbSevenZip.INT_PTR)
    Const FUNC_NAME     As String = "IVszOutStream_Seek"

    pvInitVbRuntime
    #If Logging Then
        DebugOutput "offset=" & offset_ * 10000@ & ", seekOrigin=" & seekOrigin, FUNC_NAME
    #End If
    frSeek offset_, seekOrigin, newPosition
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

Private Sub IVszOutStream_SetSize(ByVal newSize As Currency)
    Const FUNC_NAME     As String = "IVszOutStream_SetSize"

    pvInitVbRuntime
    #If Logging Then
        DebugOutput "newSize=" & newSize * 10000@, FUNC_NAME
    #End If
    frSetSize newSize
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

'=========================================================================
' IVszOutStreamFlush interface
'=========================================================================

Private Sub IVszOutStreamFlush_Flush()
    Const FUNC_NAME     As String = "IVszOutStreamFlush_Flush"

    pvInitVbRuntime
    #If Logging Then
        DebugOutput "enter", FUNC_NAME
    #End If
    Call FlushFileBuffers(m_hFile)
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

'=========================================================================
' IVszStreamGetSize interface
'=========================================================================

Private Sub IVszStreamGetSize_GetSize(ByVal size_ As VbSevenZip.INT_PTR)
    Const FUNC_NAME     As String = "IVszStreamGetSize_GetSize"
    Dim cCurrent        As Currency

    pvInitVbRuntime
    #If Logging Then
        DebugOutput "enter", FUNC_NAME
    #End If
    If VarPtr(size_) <> 0 Then
        cCurrent = frSeek(0, FILE_CURRENT)
        #If Logging Then
            DebugOutput "retval=" & frSeek(0, FILE_END, size_) * 10000@, FUNC_NAME
        #Else
            frSeek 0, FILE_END, size_
        #End If
        frSeek cCurrent, FILE_BEGIN
    End If
    #If Logging Then
        DebugOutput "exit", FUNC_NAME
    #End If
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_hFile = INVALID_HANDLE_VALUE
End Sub

Private Sub Class_Terminate()
    Const FUNC_NAME     As String = "Class_Terminate"
    
    pvInitVbRuntime
    #If Logging Then
        DebugOutput m_sFileName, MODULE_NAME & "." & FUNC_NAME
    #End If
    pvTerminate
    #If Logging Then
        DebugOutput "exit", MODULE_NAME & "." & FUNC_NAME
    #End If
End Sub

