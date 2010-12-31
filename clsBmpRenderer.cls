VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBmpRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
'   VB Gif Library Project
'   Copyright (c) 2003 Vlad Vissoultchev
'
'   Renders frames from cGifReader to 24bpp StdPictures
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "cBmpRenderer"

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private Type BITMAPINFOHEADER '40 bytes
    biSize              As Long
    biWidth             As Long
    biHeight            As Long
    biPlanes            As Integer
    biBitCount          As Integer
    biCompression       As Long
    biSizeImage         As Long
    biXPelsPerMeter     As Long
    biYPelsPerMeter     As Long
    biClrUsed           As Long
    biClrImportant      As Long
End Type

Private Type PICTDESC
    Size                As Long
    Type                As Long
    hBmpOrIcon          As Long
    hPal                As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const TRANSPARENT_COLOR1        As Long = &HCC
Private Const TRANSPARENT_COLOR2        As Long = &HFF
Private Const TRANSPARENT_SQUARE_SIZE   As Long = 8

Private m_oReader                   As clsGifReader
'--- for screen and image arrays
Private m_aScreen()                 As Byte
Private m_aCurrentScreen()          As Byte
Private m_aBackgroundScreen()       As Byte
Private m_lScreenWidth              As Long
Private m_aImageBits()              As Byte
Private m_aImageLut()               As Byte
'--- look-up table 'powers-of-two'
Private m_aPOT(-1 To 31)            As Long

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

Property Get Reader() As clsGifReader
    Set Reader = m_oReader
End Property

Property Get ScreenWidth() As Long
    ScreenWidth = m_oReader.ScreenWidth
End Property

Property Get ScreenHeight() As Long
    ScreenHeight = m_oReader.ScreenHeight
End Property

Property Get Image() As StdPicture
    Const FUNC_NAME     As String = "Image [get]"
    Dim uHead           As BITMAPINFOHEADER
    Dim hDC             As Long
    Dim hBmp            As Long
    Dim hOldBmp         As Long
    Dim lpPictDesc      As PICTDESC
    Dim aGuid(0 To 3)   As Long
    Dim hSrcDC          As Long
    
    On Error GoTo EH
    '--- prepare BITMAPINFO
    With uHead
        .biSize = Len(uHead)
        .biPlanes = 1
        .biBitCount = 24
        .biWidth = ScreenWidth
        .biHeight = -ScreenHeight
    End With
    '--- prepare dc
    hSrcDC = GetDC(0)
    hDC = CreateCompatibleDC(hSrcDC)
    hBmp = CreateCompatibleBitmap(hSrcDC, ScreenWidth, ScreenHeight)
    Call ReleaseDC(0, hSrcDC)
    '--- paint
    hOldBmp = SelectObject(hDC, hBmp)
    Call StretchDIBits(hDC, _
            0, 0, ScreenWidth, ScreenHeight, _
            0, 0, ScreenWidth, ScreenHeight, _
            m_aCurrentScreen(0), uHead, 0, vbSrcCopy)
    Call SelectObject(hDC, hOldBmp)
    '--- cleanup
    Call DeleteDC(hDC)
    '--- fill struct
    With lpPictDesc
        .Size = Len(lpPictDesc)
        .Type = vbPicTypeBitmap
        .hBmpOrIcon = hBmp
        .hPal = 0
    End With
    '--- fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGuid(0) = &H7BF80980
    aGuid(1) = &H101ABF32
    aGuid(2) = &HAA00BB8B
    aGuid(3) = &HAB0C3000
    '--- create picture from bitmap handle
    Call OleCreatePictureIndirect(lpPictDesc, aGuid(0), True, Image)
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Property Get EOF() As Boolean
    EOF = m_oReader.EOF
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(oReader As clsGifReader)
    Const FUNC_NAME     As String = "Init"
    Dim lIdx            As Long
    Dim bColor          As Byte
    
    On Error GoTo EH
    Set m_oReader = oReader
    '--- prepare screen buffer
    m_lScreenWidth = pvPadScanline(3 * ScreenWidth)
    ReDim m_aScreen(0 To m_lScreenWidth * ScreenHeight)
    '--- fill with grey&white squares pattern
    For lIdx = 0 To m_lScreenWidth - 1
        bColor = (((lIdx \ (3 * TRANSPARENT_SQUARE_SIZE)) And 1) <> 0)
        m_aScreen(lIdx) = IIf(bColor, TRANSPARENT_COLOR1, TRANSPARENT_COLOR2)
        m_aScreen(lIdx + TRANSPARENT_SQUARE_SIZE * m_lScreenWidth) = IIf(bColor, TRANSPARENT_COLOR2, TRANSPARENT_COLOR1)
    Next
    For lIdx = 0 To ScreenHeight - 1
        If ((lIdx \ TRANSPARENT_SQUARE_SIZE) And 1) = 0 Then
            CopyMemory m_aScreen(lIdx * m_lScreenWidth), m_aScreen(0), m_lScreenWidth
        Else
            CopyMemory m_aScreen(lIdx * m_lScreenWidth), m_aScreen(TRANSPARENT_SQUARE_SIZE * m_lScreenWidth), m_lScreenWidth
        End If
    Next
    '--- save background pattern for later use
    m_aBackgroundScreen = m_aScreen
    '--- success
    Init = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function MoveFirst() As Boolean
    If m_oReader.MoveFirst() Then
        m_aScreen = m_aBackgroundScreen
        MoveFirst = True
    End If
End Function

Public Function MoveNext() As Boolean
    Const FUNC_NAME     As String = "MoveNext"
    
    On Error GoTo EH
    If m_oReader.MoveNext() Then
        m_aImageBits = m_oReader.ImageBits
        m_aImageLut = m_oReader.ImageLut
        MergeCurrentImage
        MoveNext = True
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Sub MergeCurrentImage()
    Const FUNC_NAME     As String = "MergeCurrentImage"
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lPass           As Long
    Dim lStep           As Long
    Dim bRestore        As Boolean
            
    On Error Resume Next
    '--- init current screen
    m_aCurrentScreen = m_aScreen
    '--- paint on screen buffer (taking care of interlace)
    '--- warning! obfuscated code ahead :-))
    For lPass = -4 * m_oReader.IsInterlaced - 1 To (Not m_oReader.IsInterlaced) Step -1
        lStep = m_aPOT(lPass + 1)
        If lStep > 8 Then
            lStep = 8
        End If
        For lIdx = (m_aPOT(lPass) Mod 8) To m_oReader.ImageHeight - 1 Step lStep
            pvCopyPixels 3 * m_oReader.ImageLeft + (m_oReader.ImageTop + lIdx) * m_lScreenWidth, _
                    lJdx * m_oReader.ImageWidth, m_oReader.ImageWidth
            lJdx = lJdx + 1
        Next
    Next
    Select Case m_oReader.DisposalMethod
    Case ucsDsmRestorePrevious '--- do nothing (keep previous)
    Case ucsDsmRestoreBackground
        For lIdx = 0 To m_oReader.ImageHeight - 1
            lJdx = 3 * m_oReader.ImageLeft + (m_oReader.ImageTop + lIdx) * m_lScreenWidth
            CopyMemory m_aScreen(lJdx), m_aBackgroundScreen(lJdx), 3 * m_oReader.ImageWidth
        Next
    Case Else
        m_aScreen = m_aCurrentScreen
    End Select
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'--- ToDo: implement in assembly
Private Sub pvCopyPixels(ByVal lDst As Long, ByVal lSrc As Long, ByVal lLength As Long)
    Dim lTemp           As Long
    Dim lTransparent    As Long
    
    If m_oReader.IsTransparent Then
        lTransparent = m_oReader.TransparentIndex
    Else
        lTransparent = -1
    End If
    For lSrc = lSrc To lSrc + lLength - 1
        lTemp = m_aImageBits(lSrc)
        If lTemp <> lTransparent Then
            m_aCurrentScreen(lDst + 2) = m_aImageLut(3 * lTemp + 0)
            m_aCurrentScreen(lDst + 1) = m_aImageLut(3 * lTemp + 1)
            m_aCurrentScreen(lDst + 0) = m_aImageLut(3 * lTemp + 2)
        End If
        lDst = lDst + 3
    Next
End Sub

Private Function pvPadScanline(ByVal lOffset As Long)
'--- DIB section horizontal scanline padding to dword
    pvPadScanline = (lOffset + 3) And (Not 3)
End Function

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

