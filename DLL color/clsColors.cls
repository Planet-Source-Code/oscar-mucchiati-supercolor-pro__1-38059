VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    
Const RGBMAX As Byte = 255
Const YUVMAX As Byte = 255

Private iR As Byte, iG As Byte, iB As Byte
Private icY As Integer, icU As Integer, icV As Integer
Private inH As Integer, inS As Integer, inV As Integer, inL As Integer

Private iRGB_long As Long

Public Enum ColorComponentRGB
    cRed
    cGreen
    cBlu
End Enum

Public Enum ColorComponentYUV
    comp_Y
    comp_U
    comp_V
End Enum

Public Enum ColorComponentHSV
    comp_H
    comp_S
    comp_V
End Enum
Public Property Get RGB_long() As Long
    iRGB_long = RGB(iR, iG, iB)
    RGB_long = iRGB_long
End Property

Public Property Let RGB_long(ByVal lngNewValue As Long)
    If lngNewValue < 0& Or lngNewValue > 16777215 Then Exit Property
    
    iRGB_long = lngNewValue
        iR = GetCompRGB(cRed)
        iG = GetCompRGB(cGreen)
        iB = GetCompRGB(cBlu)
        
    RGB_2_ALL
        
End Property

Private Sub HSV2RGB()

    If inH = -1 Or inS = -1 Or inV = -1 Then Exit Sub

    Dim cH As Single, cS As Single, cV As Single
    Dim cR As Single, cG As Single, cB As Single
    Dim cF As Single, cI As Byte
    Dim cP As Single, cQ As Single, cT As Single
    
    cH = inH / 360
    cS = inS / 100
    cV = inV / 100
    
    If (cS = 0) Then
        cR = cV: cG = cV: cB = cV
    Else
        cH = 6 * cH
        cI = Fix(cH)
        cF = cH - cI
        cP = cV * (1 - cS)
        cQ = cV * (1 - (cS * cF))
        cT = cV * (1 - (cS * (1 - cF)))
        Select Case cI
            Case 0: cR = cV: cG = cT: cB = cP
            Case 1: cR = cQ: cG = cV: cB = cP
            Case 2: cR = cP: cG = cV: cB = cT
            Case 3: cR = cP: cG = cQ: cB = cV
            Case 4: cR = cT: cG = cP: cB = cV
            Case 5: cR = cV: cG = cP: cB = cQ
        End Select
    End If
        
    iR = cR * RGBMAX
    iG = cG * RGBMAX
    iB = cB * RGBMAX
    
End Sub

Private Function LimitsMinMax(ByVal currVal, ByVal vMin, ByVal vMax)
    If currVal < vMin Then currVal = vMin
    If currVal > vMax Then currVal = vMax
    LimitsMinMax = currVal
End Function

Private Function Max(ByVal value1, ByVal value2)
    If value1 < value2 Then Max = value2 Else Max = value1
End Function
Public Property Get WebColor() As String
    WebColor = "#" & HHex(iR) & HHex(iG) & HHex(iB)
End Property
Public Property Let WebColor(ByVal newValue As String)
    newValue = Format(Trim(Replace(newValue, "#", "")), "000000")
    
    iR = Val("&h" & Left(newValue, 2))
    iG = Val("&h" & Mid(newValue, 3, 2))
    iB = Val("&h" & Right(newValue, 2))
    
    RGB_2_ALL
    
End Property

Private Function HHex(ByVal inValue As Integer) As String
    
    Dim sStr As String
    
    sStr = Hex(inValue): If Len(sStr) < 2 Then sStr = "0" & sStr
    HHex = sStr
    
End Function
Public Function GetCompRGB(ByVal eComponent As ColorComponentRGB) As Byte
    
    GetCompRGB = CByte(iRGB_long \ 256 ^ (eComponent) And 255)

End Function
Public Property Get Red() As Byte
    Red = iR
End Property

Public Property Let Red(ByVal newValue As Byte)
    iR = newValue
        iRGB_long = RGB(iR, iB, iG)
        RGB2HSV
        RGB2YUV
End Property

Public Property Get Green() As Byte
    Green = iG
End Property

Public Property Let Green(ByVal newValue As Byte)
    iG = newValue
        iRGB_long = RGB(iR, iB, iG)
        RGB2HSV
        RGB2YUV
End Property

Public Property Get Blu() As Byte
    Blu = iB
End Property

Public Property Let Blu(ByVal newValue As Byte)
    iB = newValue
        iRGB_long = RGB(iR, iB, iG)
        RGB2HSV
        RGB2YUV
End Property

Private Sub RGB2YUV()
    
    Dim cY As Single, cU As Single, cV As Single
    
    'Modified floating-point equations
    cY = 0.299 * iR + 0.587 * iG + 0.114 * iB
    cU = 0.492 * (iB - cY)
    cV = 0.877 * (iR - cY)
    
    icY = Fix(cY)
    icU = Fix(cU)
    icV = Fix(cV)

End Sub

Private Sub YUV2RGB()
    
    Dim cR As Single, cG As Single, cB As Single
    
    cR = (icV / 0.877) + icY
    cG = 1.174 * icY - 0.682 * icV - 0.464 * icU
    cB = (icU / 0.492) + icY
        
    iR = Fix(cR)
    iG = Fix(cG)
    iB = Fix(cB)

End Sub
Public Property Get H() As Integer
    H = inH
End Property

Public Property Let H(ByVal vNewValue As Integer)
    inH = LimitsMinMax(vNewValue, 0, 360)
        HSV2RGB
        RGB2YUV
End Property

Public Property Get S() As Integer
    S = inS
End Property

Public Property Let S(ByVal vNewValue As Integer)
    inS = LimitsMinMax(vNewValue, 0, 100)
        HSV2RGB
        RGB2YUV
End Property

Public Property Get V() As Integer
    V = inV
End Property

Public Property Let V(ByVal vNewValue As Integer)
    inV = LimitsMinMax(vNewValue, 0, 100)
        HSV2RGB
        RGB2YUV
End Property

Public Property Get cY() As Integer
    cY = icY
End Property

Public Property Let cY(ByVal vNewValue As Integer)
    YUV2RGB
    RGB2HSV
End Property

Public Property Get cU() As Integer
    cU = icU
End Property

Public Property Let cU(ByVal vNewValue As Integer)
    YUV2RGB
    RGB2HSV
End Property

Public Property Get cV() As Integer
    cV = icV
End Property

Public Property Let cV(ByVal vNewValue As Integer)
    YUV2RGB
    RGB2HSV
End Property

Public Function Min(ByVal value1, ByVal value2)
    If value1 > value2 Then Min = value2 Else Min = value1
End Function

Private Sub RGB2HSV()

    Dim minRGB As Single, delta As Single
    Dim cH As Single, cS As Single, cV As Single
    Dim cR As Single, cG As Single, cB As Single
    
    cR = iR / RGBMAX
    cG = iG / RGBMAX
    cB = iB / RGBMAX
    
    minRGB = Min(cR, Min(cG, cB))
    cV = Max(cR, Max(cG, cB))
    
    delta = cV - minRGB
    
    If (cV = 0) Then cS = 0 Else cS = delta / cV
    If (cS = 0) Then
        cH = 0
    Else
        If (cR = cV) Then
            cH = 0.16666666 * (cG - cB) / delta
        Else
            If (cG = cV) Then
                cH = 0.33333333 + 0.16666666 * (cB - cR) / delta
            Else
                cH = 0.66666666 + 0.16666666 * (cR - cG) / delta
            End If
        End If
    End If
    
    If (cH < 0) Then cH = cH + 1
    
    inH = cH * 360
    inS = cS * 100
    inV = cV * 100

End Sub

Private Sub RGB_2_ALL()
    
    RGB2HSV
    RGB2YUV
    
End Sub

Private Sub Class_Initialize()

    icY = 0: icU = 0: icV = 0
    inH = 0: inS = 0: inV = 0: inL = 0
    
End Sub



Public Property Get VBColor() As String
    VBColor = "&H00" & strAlign(Hex(RGB(iR, iG, iB)), 6)
End Property

Public Property Let VBColor(ByVal vNewValue As String)

End Property

Public Function strAlign(ByVal sStr As String, ByVal iMin As Integer, Optional ByVal sCh As String = "0") As String
    
    sStr = Trim(sStr)
        If Len(sStr) < iMin Then sStr = String(iMin - Len(sStr), sCh) & sStr
    
    strAlign = sStr
    
End Function

Public Property Get HexColor() As String
    HexColor = Hex(RGB(iR, iG, iB))
End Property

Public Property Let HexColor(ByVal vNewValue As String)

End Property

Public Function Mix(ByRef newColor As clsColors) As clsColors

End Function
