Attribute VB_Name = "modify_resolution"
Option Explicit

Global Xpix As Long
Global Ypix As Long
Global lColorQuality As Long
Dim dx As New DirectX7
Dim dd As DirectDraw7


Function SetResolution(resX As Long, resY As Long, Bbp As Long, Optional refresh As Long = 0)
Set dd = dx.DirectDrawCreate("") 'create the direct draw object
dd.SetDisplayMode resX, resY, Bbp, refresh, DDSDM_DEFAULT 'set the display mode
End Function

Function ResetResolution()
dd.RestoreDisplayMode
End Function

Function GetSupportedResolutionIndex(index As Integer, section As String) As Long
Dim ddsd As DDSURFACEDESC2
Dim DisplayModesEnum As DirectDrawEnumModes
Dim i As Integer
Set dd = dx.DirectDrawCreate("") 'create the direct draw object
Set DisplayModesEnum = dd.GetDisplayModesEnum(0, ddsd) 'create enum object
DisplayModesEnum.GetItem index, ddsd 'the the supported resolution
Select Case section
Case "X" 'if user wants the width
GetSupportedResolutionIndex = ddsd.lWidth
Case "Y" 'the height
GetSupportedResolutionIndex = ddsd.lHeight
Case "B" 'the color depth
GetSupportedResolutionIndex = ddsd.ddpfPixelFormat.lRGBBitCount
Case "R" 'the refresh rate
GetSupportedResolutionIndex = ddsd.lRefreshRate
End Select
End Function

Function GetSupportedResolutionCount() As Long
Dim ddsd As DDSURFACEDESC2
Dim DisplayModesEnum As DirectDrawEnumModes
Dim i As Integer
Set dd = dx.DirectDrawCreate("") 'create the direct draw object
Set DisplayModesEnum = dd.GetDisplayModesEnum(0, ddsd) 'create enum object
GetSupportedResolutionCount = DisplayModesEnum.GetCount 'get the count
End Function



