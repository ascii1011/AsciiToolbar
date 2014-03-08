Attribute VB_Name = "display_info"


Sub prcDetermineResolution(intWidth As Long, intHeight As Long)
    intWidth = Screen.Width \ Screen.TwipsPerPixelX
    intHeight = Screen.Height \ Screen.TwipsPerPixelY
End Sub
