VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGifReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
'   VB Gif Library Project
'   Copyright (c) 2003 Vlad Vissoultchev
'
'   GIF87a/89a reader. Implements an LZW decoder. Warning! use of this
'     code in commercial applications may fall under patent claims
'     from Unisys which are holding patents on LZW algorithm.
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "cGifReader"

'=========================================================================
' Events
'=========================================================================

Event Progress(ByVal CurrentLine As Long)
Event ImageComplete()

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const ERR_INVALID_GIF_FILE      As String = "Invalid GIF file"
Private Const ERR_UNEXPECTED_BLOCK      As String = "Unexpected image block"
Private Const ERR_PAST_END_OF_STACK     As String = "Past end of stack (stacksize=4000)"
Private Const ERR_INVALID_LZW_CODE      As String = "Invalid LZW code encountered"
Private Const ERR_INPUT_PAST_EOF        As String = "Input past end of file"
Private Const STR_GIF87A                As String = "GIF87a"
Private Const STR_GIF89A                As String = "GIF89a"
Private Const MAX_BITS                  As Long = 12
Private Const TABLE_SIZE                As Long = 2 ^ MAX_BITS

'--- look-up table 'powers-of-two'
Private m_aPOT(-1 To 31)            As Long
'--- GIF file
Private m_nFile                     As Integer
Private m_sFileName                 As String
Private m_lFirstFrameLoc            As Long
Private m_bEOF                      As Boolean
'--- for GIF content
Private m_uHeader                   As UcsGifHeader
Private m_uImageDesc                As UcsGifImageDescriptor
Private m_uGraphicControl           As UcsGifGraphicControl
Private m_lFrameIndex               As Long
'--- for current frame buffers
Private m_aImageBits()              As Byte
Private m_aGlobalLut(0 To 767)      As Byte '--- 767 = 3 * 256 - 1
Private m_aImageLut(0 To 767)       As Byte
'--- for LZW decoder
Private m_lInitBits                 As Long
Private m_lClearTable               As Long
Private m_lInputBitCount            As Long
Private m_lInputBitBuffer           As Long
Private m_lCurrentBits              As Long
Private m_lMaxCode                  As Long
Private m_lSubBlockSize             As Long
Private m_aPrefixCode(0 To TABLE_SIZE) As Long
Private m_aAppendChar(0 To TABLE_SIZE) As Byte

Private Type UcsGifHeader
    aSigVersion(0 To 5)     As Byte
    nScreenWidth            As Integer
    nScreenHeight           As Integer
    bFlags                  As Byte
    bBackgroungColor        As Byte
    bAspectRatio            As Byte
End Type

Private Type UcsGifImageDescriptor
    nImageLeft              As Integer
    nImageTop               As Integer
    nImageWidth             As Integer
    nImageHeight            As Integer
    bFlags                  As Byte
End Type

Private Type UcsGifGraphicControl
    cbSize                  As Byte
    bFlags                  As Byte
    nDelayTime              As Integer
    bTransparentColor       As Byte
    bTerminator             As Byte
End Type

Private Enum UcsGifFlags
    '--- for header flags
    ucsGflGlobalLut = &H80
    ucsGflColorResolution = &H70
    ucsGflGlobalLutSorted = &H4
    ucsGflGlobalLutSize = &H7
    '--- for image descriptor flags
    ucsGflLocalLut = &H80
    ucsGflInterlace = &H40
    ucsGflLocalLutSorted = &H20
    ucsGflLocalLutSize = &H7
    '--- for graphics control
    ucsGflDisposalMethod = &H1C
    ucsGflUserInput = &H2
    ucsGflTransparentColor = &H1
End Enum

Private Enum UcsGifFileBlock
    ucsGblImageBlock = &H2C         '--- ","
    ucsGblExtension = &H21          '--- "!"
    ucsGblTrailer = &H3B            '--- ";"
End Enum

Private Enum UcsGifExtensionType
    ucsGexGraphicsControl = &HF9    '--- transparency etc. extension
End Enum

Public Enum UcsGifDisposalMethod
    ucsDsmNotSpecified
    ucsDsmDontDispose
    ucsDsmRestoreBackground
    ucsDsmRestorePrevious
End Enum

'=========================================================================
' Error management
'=========================================================================

Private Sub RaiseError(sFunction As String)
    With Err
        .Raise .Number, MODULE_NAME & "." & sFunction & IIf(Erl <> 0, "(" & Erl & ")", "") & vbCrLf _
                & .Source, .Description, .HelpFile, .HelpContext
    End With
End Sub

Private Sub PrintError(sFunction As String)
    Debug.Print MODULE_NAME; "."; sFunction; IIf(Erl <> 0, "(" & Erl & ")", ""); ": "; Err.Description
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get FileName() As String
    FileName = m_sFileName
End Property

Property Get SigVersion() As String
    SigVersion = StrConv(m_uHeader.aSigVersion, vbUnicode)
End Property

Property Get HasGlobalLut() As Boolean
    HasGlobalLut = (pvGetFlag(m_uHeader.bFlags, ucsGflGlobalLut) <> 0)
End Property

Property Get ScreenWidth() As Long
    ScreenWidth = m_uHeader.nScreenWidth
End Property

Property Get ScreenHeight() As Long
    ScreenHeight = m_uHeader.nScreenHeight
End Property

Property Get BackgroundColor() As Long
    BackgroundColor = RGB(m_aImageLut(3 * m_uHeader.bBackgroungColor), _
            m_aImageLut(3 * m_uHeader.bBackgroungColor + 1), _
            m_aImageLut(3 * m_uHeader.bBackgroungColor + 2))
End Property

Property Get BackgroundIndex() As Long
    BackgroundIndex = m_uHeader.bBackgroungColor
End Property

Property Get GlobalLutSize() As Long
    GlobalLutSize = m_aPOT(1 + pvGetFlag(m_uHeader.bFlags, ucsGflGlobalLutSize))
End Property

Property Get IsInterlaced() As Boolean
    IsInterlaced = (pvGetFlag(m_uImageDesc.bFlags, ucsGflInterlace) <> 0)
End Property

Property Get HasLocalLut() As Boolean
    HasLocalLut = (pvGetFlag(m_uImageDesc.bFlags, ucsGflLocalLut) <> 0)
End Property

Property Get LocalLutSize() As Long
    LocalLutSize = m_aPOT(1 + pvGetFlag(m_uImageDesc.bFlags, ucsGflLocalLutSize))
End Property

Property Get ImageLeft() As Long
    ImageLeft = m_uImageDesc.nImageLeft
End Property

Property Get ImageTop() As Long
    ImageTop = m_uImageDesc.nImageTop
End Property

Property Get ImageWidth() As Long
    ImageWidth = m_uImageDesc.nImageWidth
End Property

Property Get ImageHeight() As Long
    ImageHeight = m_uImageDesc.nImageHeight
End Property

Property Get IsTransparent() As Boolean
    IsTransparent = (pvGetFlag(m_uGraphicControl.bFlags, ucsGflTransparentColor) <> 0)
End Property

Property Get TransparentColor() As Long
    TransparentColor = RGB(m_aImageLut(3 * m_uGraphicControl.bTransparentColor), _
            m_aImageLut(3 * m_uGraphicControl.bTransparentColor + 1), _
            m_aImageLut(3 * m_uGraphicControl.bTransparentColor + 2))
End Property

Property Get TransparentIndex() As Long
    TransparentIndex = m_uGraphicControl.bTransparentColor
End Property

'--- note: usually this is interpreted as 'gif animation is looped'
Property Get UserInput() As Boolean
    UserInput = (pvGetFlag(m_uGraphicControl.bFlags, ucsGflUserInput) <> 0)
End Property

Property Get DisposalMethod() As UcsGifDisposalMethod
    DisposalMethod = pvGetFlag(m_uGraphicControl.bFlags, ucsGflDisposalMethod)
End Property

Property Get ImageLut() As Byte()
    ImageLut = m_aImageLut
End Property

Property Get ImageBits() As Byte()
    ImageBits = m_aImageBits
End Property

Property Get FrameIndex() As Long
    FrameIndex = m_lFrameIndex
End Property

Property Get EOF() As Boolean
    EOF = m_bEOF
End Property

Property Get DelayTime() As Long
    DelayTime = m_uGraphicControl.nDelayTime
End Property

'=========================================================================
' Methods
'=========================================================================

Private Function pvFileExists(sFileName As String) As Boolean
    On Error Resume Next
    pvFileExists = (GetAttr(sFileName) <> -1)
    On Error GoTo 0
End Function

Public Function Init(sFileName As String) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim lIdx            As Long
    Dim bColor          As Boolean
    
    On Error GoTo EH
    '--- init/clear member vars
    m_sFileName = sFileName
    FillMemory m_uImageDesc, Len(m_uImageDesc), 0
    FillMemory m_uGraphicControl, Len(m_uGraphicControl), 0
    m_lFrameIndex = -1
    m_bEOF = False
    '--- check if file exists
    If Not pvFileExists(sFileName) Then
        Exit Function
    End If
    '--- open file (first close previous)
    pvCloseFile
    m_nFile = FreeFile()
    Open FileName For Binary As #m_nFile
    '--- get file header
    pvReadBuffer VarPtr(m_uHeader), Len(m_uHeader) '--- 13
    If SigVersion <> STR_GIF87A And SigVersion <> STR_GIF89A Then
        Err.Raise vbObject, , ERR_INVALID_GIF_FILE
    End If
    '--- get global LUT
    If HasGlobalLut Then
        pvReadBuffer VarPtr(m_aGlobalLut(0)), 3 * GlobalLutSize
    End If
    m_lFirstFrameLoc = Seek(m_nFile)
    '--- success
    Init = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function MoveNext() As Boolean
    Const FUNC_NAME     As String = "MoveNext"
   
    On Error GoTo EH
    '--- check if anything's left
    If m_nFile = 0 Or m_bEOF Then
        Exit Function
    End If
    Do While True
        Select Case pvReadByte()
        Case ucsGblImageBlock
            m_lFrameIndex = m_lFrameIndex + 1
            '--- get image desc
            pvReadBuffer VarPtr(m_uImageDesc), Len(m_uImageDesc) '--- 9
            '--- get image LUT
            If HasLocalLut Then
                pvReadBuffer VarPtr(m_aImageLut(0)), 3 * LocalLutSize
            Else
                CopyMemory m_aImageLut(0), m_aGlobalLut(0), 3 * 256
            End If
            '--- init vars for LZW decoding
            m_lInitBits = pvReadByte() + 1
            m_lClearTable = m_aPOT(m_lInitBits - 1)
            m_lInputBitCount = 0
            m_lInputBitBuffer = 0
            m_lCurrentBits = m_lInitBits
            m_lMaxCode = m_aPOT(m_lCurrentBits) - 1
            m_lSubBlockSize = 0
            ReDim m_aImageBits(0 To ImageWidth * ImageHeight)
            pvLzwExpand m_aImageBits, UBound(m_aImageBits)
            '--- read to the end of block
            Do
                '--- skip to the end end of sub-block
                Do While m_lSubBlockSize > 0
                    pvReadSubBlockByte
                Loop
                '--- check for block terminator
                m_lSubBlockSize = pvReadByte()
            Loop While m_lSubBlockSize > 0
            RaiseEvent ImageComplete
            Exit Do
        Case ucsGblExtension
            '--- look for 'Graphic Control Label' extension
            Select Case pvReadByte()
            Case ucsGexGraphicsControl
                '--- fill member struct
                pvReadBuffer VarPtr(m_uGraphicControl), Len(m_uGraphicControl)
            Case Else
                '--- unknown extension
                pvSkipBlock
            End Select
        Case ucsGblTrailer
            m_bEOF = True
            Exit Function
        Case 0 '--- silence this just in case
            Debug.Print MODULE_NAME; "."; FUNC_NAME; ": "; ERR_UNEXPECTED_BLOCK; " = 0"
        Case Else
            Err.Raise vbObjectError + 1, , ERR_UNEXPECTED_BLOCK
        End Select
    Loop
    '--- success
    MoveNext = True
    Exit Function
EH:
    m_bEOF = True
    RaiseError FUNC_NAME
End Function

Public Function MoveFirst() As Boolean
    Const FUNC_NAME     As String = "MoveFirst"
    
    On Error GoTo EH
    '--- state check
    If m_nFile = 0 Then
        Exit Function
    End If
    Seek #m_nFile, m_lFirstFrameLoc
    m_lFrameIndex = -1
    m_bEOF = False
    '--- success
    MoveFirst = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function MoveLast() As Long
    Const FUNC_NAME     As String = "MoveLast"
    
    On Error GoTo EH
    '--- state check
    If m_nFile = 0 Then
        Exit Function
    End If
    If Not EOF Then
        Do While True
            Select Case pvReadByte()
            Case ucsGblImageBlock
                m_lFrameIndex = m_lFrameIndex + 1
                pvReadBuffer VarPtr(m_uImageDesc), Len(m_uImageDesc) '--- 9
                If HasLocalLut Then
                    pvReadBuffer VarPtr(m_aImageLut(0)), 3 * LocalLutSize
                End If
                pvReadByte '--- initial bits
                pvSkipBlock
            Case ucsGblExtension
                pvReadByte '--- extension type
                pvSkipBlock
            Case ucsGblTrailer
                m_bEOF = True
                Exit Do
            Case 0
            Case Else
                Err.Raise vbObjectError + 1, , ERR_UNEXPECTED_BLOCK
            End Select
        Loop
    End If
    '--- success
    MoveLast = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'= private ===============================================================

Private Function pvGetFlag(ByVal lFlags As Long, ByVal lMask As UcsGifFlags) As Long
    If lMask > 0 Then
        pvGetFlag = (lFlags And lMask)
        Do While (lMask And 1) = 0
            lMask = lMask \ 2
            pvGetFlag = pvGetFlag \ 2
        Loop
    End If
End Function

Private Function pvPadScanline(ByVal lOffset As Long)
'--- DIB section horizontal scanline padding to dword
    pvPadScanline = (lOffset + 3) And (Not 3)
End Function

'= I/O handling ==========================================================

Private Function pvReadByte() As Byte
    Const FUNC_NAME     As String = "pvReadByte"
    
    On Error GoTo EH
    If VBA.EOF(m_nFile) Then
        Err.Raise vbObjectError + 4, , ERR_INPUT_PAST_EOF
    End If
    Get #m_nFile, , pvReadByte
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvReadBuffer(ByVal pAddr As Long, ByVal lSize)
    Const FUNC_NAME     As String = "pvReadBuffer"
    Dim lIdx            As Long
    
    On Error GoTo EH
    '--- read from stream to local buffer
    ReDim aBuf(0 To lSize) As Byte
    For lIdx = 0 To lSize - 1
        aBuf(lIdx) = pvReadByte()
    Next
    '--- copy if necessary
    If pAddr <> 0 Then
        CopyMemory ByVal pAddr, aBuf(0), lSize
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function pvReadSubBlockByte() As Byte
    Const FUNC_NAME     As String = "pvReadSubBlockByte"
    
    On Error GoTo EH
    If m_lSubBlockSize <= 0 Then
        m_lSubBlockSize = pvReadByte()
        '--- workaround for 3D Studio R4's non-compliant GIFs
        If m_lSubBlockSize = 0 Then
            m_lSubBlockSize = 256
        End If
    End If
    pvReadSubBlockByte = pvReadByte()
    m_lSubBlockSize = m_lSubBlockSize - 1
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvCloseFile()
    Const FUNC_NAME     As String = "pvCloseFile"
    
    On Error GoTo EH
    If m_nFile <> 0 Then
        Close #m_nFile
        m_nFile = 0
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvSkipBlock()
    Do
        m_lSubBlockSize = pvReadByte()
        Seek #m_nFile, Seek(m_nFile) + m_lSubBlockSize
    Loop While m_lSubBlockSize > 0
End Sub

'= LZW decompressor ======================================================

Private Function pvLzwReadCode() As Long
    Const FUNC_NAME     As String = "pvLzwReadCode"
    
    On Error GoTo EH
    Do While m_lInputBitCount < m_lCurrentBits
        m_lInputBitBuffer = m_lInputBitBuffer Or (pvReadSubBlockByte() * m_aPOT(m_lInputBitCount))
        m_lInputBitCount = m_lInputBitCount + 8
    Loop
    pvLzwReadCode = m_lInputBitBuffer And (m_aPOT(m_lCurrentBits) - 1)
    m_lInputBitBuffer = m_lInputBitBuffer \ m_aPOT(m_lCurrentBits)
    m_lInputBitCount = m_lInputBitCount - m_lCurrentBits
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvLzwDecodeString(aStack() As Byte, ByVal lIdx As Long, _
            ByVal lCode As Long) As Long
    Const FUNC_NAME     As String = "pvLzwDecodeString"
    
    On Error GoTo EH
    Do While lCode >= m_lClearTable
        aStack(lIdx) = m_aAppendChar(lCode)
        lIdx = lIdx + 1
        lCode = m_aPrefixCode(lCode)
        If lIdx > UBound(aStack) Then
            Err.Raise vbObjectError + 2, , ERR_PAST_END_OF_STACK
        End If
    Loop
    aStack(lIdx) = lCode
    pvLzwDecodeString = lIdx
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvLzwExpand(aBuffer() As Byte, ByVal lBufSize)
    Const FUNC_NAME     As String = "pvLzwExpand"
    Dim lIdx            As Long
    Dim lNewCode        As Long
    Dim lOldCode        As Long
    Dim lNextCode       As Long
    Dim bCharacter      As Byte
    Dim bClearFlag      As Boolean
    Dim aStack(0 To 4000) As Byte
    Dim lStackIdx       As Long
    Dim lPrevProgess    As Long
    
    On Error GoTo EH
    lNextCode = m_lClearTable + 2 '--- first code = m_lClearTable + 2
    bClearFlag = True
    lNewCode = pvLzwReadCode()
    Do While lIdx < lBufSize
        lNewCode = pvLzwReadCode()
        '--- check for terminator
        If lNewCode = m_lClearTable + 1 Then '--- terminator = m_lClearTable + 1
            Exit Sub
        End If
        If bClearFlag Then
            bClearFlag = False
            lOldCode = lNewCode
            bCharacter = lNewCode
            aBuffer(lIdx) = bCharacter
            lIdx = lIdx + 1
        ElseIf lNewCode = m_lClearTable Then
            bClearFlag = True
            m_lCurrentBits = m_lInitBits
            m_lMaxCode = m_aPOT(m_lCurrentBits) - 1
            lNextCode = m_lClearTable + 2 '--- first code = m_lClearTable + 2
        Else
            '--- decode string
            If lNewCode < lNextCode Then
                lStackIdx = pvLzwDecodeString(aStack, 0, lNewCode)
            ElseIf lNewCode = lNextCode Then
                aStack(0) = bCharacter
                lStackIdx = pvLzwDecodeString(aStack, 1, lOldCode)
            Else
                Err.Raise vbObjectError + 3, , ERR_INVALID_LZW_CODE
            End If
            '--- save first char
            bCharacter = aStack(lStackIdx)
            '--- reverse copy stack
            Do While lStackIdx >= 0
                aBuffer(lIdx) = aStack(lStackIdx)
                lStackIdx = lStackIdx - 1
                lIdx = lIdx + 1
            Loop
            '--- keep char table up-to-date
            m_aPrefixCode(lNextCode) = lOldCode
            m_aAppendChar(lNextCode) = bCharacter
            lNextCode = lNextCode + 1
            '--- expand code bitsize if max reached
            If lNextCode > m_lMaxCode Then
                If m_lCurrentBits < MAX_BITS Then
                    m_lCurrentBits = m_lCurrentBits + 1
                    m_lMaxCode = m_aPOT(m_lCurrentBits) - 1
                End If
            End If
            lOldCode = lNewCode
        End If
        '--- report progress
        lStackIdx = lIdx \ ImageWidth
        If lStackIdx >= lPrevProgess + 10 Then
            RaiseEvent Progress(lStackIdx)
            lPrevProgess = lStackIdx
        End If
    Loop
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Dim lIdx            As Long
    
    '--- init look-up table for fast 2 ^ x
    m_aPOT(-1) = 0
    m_aPOT(0) = 1
    For lIdx = 1 To 30
        m_aPOT(lIdx) = 2 * m_aPOT(lIdx - 1)
    Next
    m_aPOT(31) = &H80000000
End Sub

Private Sub Class_Terminate()
    Dim lErrNum         As Long
    Dim sErrSrc         As String
    Dim sErrDesc        As String
    
    '--- preserve error info and try not to throw one
    '--- best practices: "destructors are never throwing an exception"
    lErrNum = Err.Number
    sErrSrc = Err.Source
    sErrDesc = Err.Description
    On Error Resume Next
    pvCloseFile
    On Error GoTo 0
    Err.Number = lErrNum
    Err.Source = sErrSrc
    Err.Description = sErrDesc
End Sub
