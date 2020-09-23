Attribute VB_Name = "mdlGradient"
Option Explicit

Public Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer 'Ushort value
    Green As Integer 'Ushort value
    Blue As Integer 'ushort value
    Alpha As Integer 'ushort
End Type
Public Type GRADIENT_RECT
    UpperLeft As Long  'In reality this is a UNSIGNED Long
    LowerRight As Long 'In reality this is a UNSIGNED Long
End Type

Public Const GRADIENT_FILL_RECT_H As Long = &H0 'In this mode, two endpoints describe a rectangle. The rectangle is
'defined to have a constant color (specified by the TRIVERTEX structure) for the left and right edges. GDI interpolates
'the color from the top to bottom edge and fills the interior.
Public Const GRADIENT_FILL_RECT_V  As Long = &H1 'In this mode, two endpoints describe a rectangle. The rectangle
' is defined to have a constant color (specified by the TRIVERTEX structure) for the top and bottom edges. GDI interpolates
' the color from the top to bottom edge and fills the interior.
Public Const GRADIENT_FILL_TRIANGLE As Long = &H2 'In this mode, an array of TRIVERTEX structures is passed to GDI
'along with a list of array indexes that describe separate triangles. GDI performs linear interpolation between triangle vertices
'and fills the interior. Drawing is done directly in 24- and 32-bpp modes. Dithering is performed in 16-, 8.4-, and 1-bpp mode.
Public Const GRADIENT_FILL_OP_FLAG As Long = &HFF

Public Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Public Function LongToUShort(Unsigned As Long) As Integer
    On Error Resume Next
    'A small function to convert from long to unsigned short
    LongToUShort = CInt(Unsigned - &H10000)
End Function
Public Sub doGradient(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    
    Dim vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT

    'from black
    With vert(0)
        .x = x
        .y = y
        .Red = LongToUShort(&HFF00&)
        .Green = LongToUShort(&HFF00&)
        .Blue = LongToUShort(&HFF00&)
        .Alpha = 0&
    End With

    'to blue
    With vert(1)
        .x = x + w
        .y = y + h
        .Red = LongToUShort(&H3000&)
        .Green = LongToUShort(&HA000&)
        .Blue = LongToUShort(&HFF00&)
        .Alpha = 0&
    End With

    gRect.UpperLeft = 0
    gRect.LowerRight = 1

    GradientFillRect hDC, vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_H
    
End Sub

