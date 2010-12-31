Attribute VB_Name = "modColourDlg"

Option Explicit

Private Const CC_RGBINIT = &H1
Private Const CC_FULLOPEN = &H2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_SHOWHELP = &H8
Private Const CC_ENABLEHOOK = &H10
Private Const CC_ENABLETEMPLATE = &H20
Private Const CC_ENABLETEMPLATEHANDLE = &H40
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_ANYCOLOR = &H100

Private Const COLOR_FLAGS = CC_FULLOPEN Or CC_ANYCOLOR Or CC_RGBINIT

Private Type CHOOSECOLORS
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORS) As Long

Function ShowColorDlg(ByVal hwnd As Long, Optional DefCol As Long = vbBlack) As Long

Dim ColorDialog As CHOOSECOLORS, customcolors() As Byte, i As Integer, ret As Long

If ColorDialog.lpCustColors = "" Then
    ReDim customcolors(0 To 16 * 4 - 1) As Byte  'resize the array
    
    For i = LBound(customcolors) To UBound(customcolors)
      customcolors(i) = 254
    Next i
    
    ColorDialog.lpCustColors = StrConv(customcolors, vbUnicode)   ' convert array
End If

ColorDialog.hwndOwner = hwnd
ColorDialog.lStructSize = Len(ColorDialog)
ColorDialog.flags = COLOR_FLAGS
ColorDialog.rgbResult = DefCol

ret = ChooseColor(ColorDialog)

If ret Then
    ShowColorDlg = ColorDialog.rgbResult
Else
    ShowColorDlg = -1
End If
End Function
