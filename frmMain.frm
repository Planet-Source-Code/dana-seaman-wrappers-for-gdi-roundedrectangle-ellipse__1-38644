VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "GDI+ Test - Button Demos"
   ClientHeight    =   5985
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "GDIPlus Test"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3360
      Top             =   -540
   End
   Begin VB.PictureBox picGrapes 
      Height          =   915
      Left            =   3660
      ScaleHeight     =   855
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   2580
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   6
      Left            =   0
      Picture         =   "frmMain.frx":AA47
      Top             =   -540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   0
      Left            =   480
      Picture         =   "frmMain.frx":B311
      Top             =   -540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   1
      Left            =   960
      Picture         =   "frmMain.frx":BBDB
      Top             =   -540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   2
      Left            =   1440
      Picture         =   "frmMain.frx":C4A5
      Top             =   -540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   3
      Left            =   1920
      Picture         =   "frmMain.frx":CD6F
      Top             =   -540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   4
      Left            =   2400
      Picture         =   "frmMain.frx":D639
      Top             =   -540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   5
      Left            =   2880
      Picture         =   "frmMain.frx":DF03
      Top             =   -540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRedraw 
      Caption         =   "&Redraw Demo"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Avery P. - 7/30/2002
' Examples are from the GDI+ portion of the Platform SDK
' D.Seaman - 28-Aug-2002
' Wrappers for Ellipse and MultiStyle
' Rectangles added to GDIPlusAPI module

Dim Stars(0 To 15)   As POINTAPI
Dim lAniStep         As Long
Dim token            As Long ' Needed to close GDI+

Private Sub Form_Load()
   Dim I As Long
   ' Load the GDI+ Dll
   Dim GpInput          As GdiplusStartupInput
   GpInput.GdiplusVersion = 1
   If GdiplusStartup(token, GpInput) <> Ok Then
      MsgBox "Error loading GDI+!", vbCritical
      Unload Me
   End If
   For I = 0 To 15
      Stars(I).x = Rnd * 120
      Stars(I).y = Rnd * 40
   Next
   Me.Show
   mnuRedraw_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' Unload the GDI+ Dll
   GdiplusShutdown (token)
End Sub

Private Sub mnuRedraw_Click()
   
   Cls   ' Clear the window

   ' Uncomment one to see it's demo
   'DrawCurves
   'DrawScaling
   'DrawSkewed
   'DrawTexturedLine
   'DrawLineCaps
   'DrawCustomDashed
   'DrawSolidShape
   
   DrawEllipticalButtons
   DrawRectangularButtons
   DrawRoundedRectangleButtons
   
   'DrawHatchShape
   'DrawThumbnail     ' You'll need to change the image filename
   'DrawHorizGradient
   'DrawDiagGradient
   'DrawPathGradient
   'BMPtoPNG
   'BMPtoJPEG
   'BMPtoJPEG_Params
   'DrawCachedBitmap  ' WARNING: Running this with AutoRedraw = True is not a good idea!
   'DrawAlphaLines
   'DrawColorMatrix
   'DrawAlphaPixels
   'DrawAntiAliasText
    'DrawRotated
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub

'======================================================================
' THE SAMPLES!
'======================================================================
Private Sub DrawEllipticalButtons()
   DrawGDIPEllipse hDC, _
      270, 10, 120, 40, _
      BrushTypeLinearGradient, _
      1, PenTypeSolidColor, _
      XPBlue, _
      White, XPGradient
      
   DrawShell32Icon hDC, 80, 278, 23, True
   
   DrawGDIPFormattedText hDC, 298, 18, 66, 24, _
      "XP Style Small Icon", _
      StringAlignmentCenter, Black
   DrawGDIPFocusRect hDC, 298, 18, 66, 24, _
      1, Black, DashStyleDot
      
   '##### Ellipse 2
   DrawGDIPEllipse hDC, _
      270, 60, 120, 40, _
      BrushTypePathGradient, _
      2, PenTypeSolidColor, _
      Yellow, _
      Yellow, OrangeRed

   DrawShell32Icon hDC, 4, 278, 73, True
   
   DrawGDIPFormattedText hDC, 298, 68, 66, 24, _
      "Elliptical Small Icon", _
      StringAlignmentCenter, Black
   DrawGDIPFocusRect hDC, 298, 68, 66, 24, _
      1, Black, DashStyleDot
      
End Sub
Private Sub DrawRectangularButtons()
   '##### RR 1
   DrawGDIPMultiStyleRectangle hDC, _
      10, 10, 130, 52, 0, _
      BrushTypeLinearGradient, _
      1, PenTypeSolidColor, XPBlue, , _
      White, XPGradient
      
   DrawShell32Icon hDC, 40, 16, 16
   
   DrawGDIPFormattedText hDC, 49, 14, 77, 32, _
      "XP Style GDI+ Button", _
      StringAlignmentCenter, Black
   DrawGDIPFocusRect hDC, 14, 14, 112, 34, _
      1, Black, DashStyleDot
   '##### RR 2
   ' Gradient Fill (XP Button colors)
  ' DrawGDIPMultiStyleRectangle hdc, _
      10, 60, 130, 100, 0, _
      BrushTypeSolidColor, _
      1, PenTypeSolidColor, White, , _
      Black, Black
   
  ' DrawShell32Icon hdc, 20, 14, 64
   
  ' DrawGDIPFormattedText hdc, 48, 62, 77, 36, _
      "Owner Draw Background", _
      StringAlignmentCenter, Black
      
   '##### RR 3
   DrawGDIPMultiStyleRectangle hDC, _
      10, 110, 130, 150, 0, _
      BrushTypeLinearGradient, _
      3, PenTypeSolidColor, &HFF80FFFF, &HFF0080FF, _
      &HFF0080FF, Cyan

   DrawShell32Icon hDC, 79, 14, 114
   
   DrawGDIPFormattedText hDC, 48, 112, 77, 36, _
      "Raised 3D Gradient Fill", _
      StringAlignmentCenter, Black
      
   '##### RR 4
   DrawGDIPMultiStyleRectangle hDC, _
      10, 160, 130, 200, 0, _
      BrushTypeLinearGradient, _
      3, PenTypeSolidColor, &HFF0080FF, &HFF80FFFF, _
      Cyan, &HFF0080FF
   DrawGDIPFormattedText hDC, 48, 162, 77, 36, _
      "Sunken 3D Gradient Fill", _
      StringAlignmentCenter, Black
   DrawShell32Icon hDC, 80, 14, 164
   
   '##### RR 5
   DrawGDIPMultiStyleRectangle hDC, _
      10, 210, 130, 250, 0, _
      BrushTypeTextureFill, _
      3, PenTypeSolidColor, &HFFFCF9C3, , _
      &HFFFCF9C3, &HFFB08218, , _
      AppPath & "Mogno.gif"
   DrawGDIPFormattedText hDC, 48, 216, 77, 28, _
      "Mahogany Fill Solid Border", _
      StringAlignmentCenter, White
   DrawGDIPFocusRect hDC, 48, 216, 77, 28, _
      1, White, DashStyleDot
   DrawShell32Icon hDC, 4, 14, 214

   '##### RR 6
   DrawGDIPMultiStyleRectangle hDC, _
      10, 260, 130, 300, 0, _
      BrushTypeTextureFill, _
      3, PenTypeSolidColor, XPGoldLight, DarkBrown, _
      XPGoldLight, XPGoldDark, , _
      AppPath & "Spruce.gif"
   DrawGDIPFormattedText hDC, 48, 266, 77, 28, _
      "Spruce Fill 3D Raised Border", _
      StringAlignmentCenter, White
   DrawGDIPFocusRect hDC, 48, 266, 77, 28, _
      1, White, DashStyleDot
   DrawShell32Icon hDC, 4, 14, 264
End Sub
Private Sub DrawRoundedRectangleButtons()
   '##### RR 1
  ' DrawGDIPMultiStyleRectangle hdc, _
      140, 10, 260, 52, 4, _
      BrushTypeLinearGradient, _
      1, PenTypeSolidColor, XPBlue, , _
      White, XPGradient
      
      
   'DrawShell32Icon hdc, 40, 146, 16
   
  ' DrawGDIPFormattedText hdc, 179, 14, 77, 32, _
      "XP Style GDI+ Button", _
      StringAlignmentCenter, Black
  ' DrawGDIPFocusRect hdc, 144, 14, 112, 34, _
      1, Black, DashStyleDot
      
   '##### RR 2
   ' Gradient Fill (XP Button colors)
   DrawGDIPMultiStyleRectangle hDC, _
      140, 60, 260, 100, 4, _
      BrushTypeLinearGradient, _
      1, PenTypeSolidColor, Black, , _
      Yellow, OrangeRed
   
   DrawShell32Icon hDC, 20, 144, 64
   
   DrawGDIPFormattedText hDC, 178, 62, 77, 36, _
      "RoundedRect Gradient Fill", _
      StringAlignmentCenter, Black
      
   '##### RR 3
   DrawGDIPMultiStyleRectangle hDC, _
      140, 110, 260, 150, 4, _
      BrushTypeLinearGradient, _
      3, PenTypeSolidColor, XPGoldLight, XPGoldDark, _
      XPGoldDark, XPGoldLight

   DrawShell32Icon hDC, 79, 144, 114
   
   DrawGDIPFormattedText hDC, 178, 112, 77, 36, _
      "Raised 3D Gradient Fill", _
      StringAlignmentCenter, Black
      
   '##### RR 4
   DrawGDIPMultiStyleRectangle hDC, _
      140, 160, 260, 200, 4, _
      BrushTypeLinearGradient, _
      3, PenTypeSolidColor, XPGoldDark, XPGoldLight, _
       XPGoldLight, XPGoldDark
      
   DrawGDIPFormattedText hDC, 178, 162, 77, 36, _
      "Sunken 3D Gradient Fill", _
      StringAlignmentCenter, Black
   DrawShell32Icon hDC, 80, 144, 164
   
   '##### RR 5
   DrawGDIPMultiStyleRectangle hDC, _
      140, 210, 260, 250, 4, _
      BrushTypeTextureFill, _
      3, PenTypeSolidColor, XPGoldLight, , _
      XPGoldLight, XPGoldDark, , _
      AppPath & "Mogno.gif"
   DrawGDIPFormattedText hDC, 178, 216, 77, 28, _
      "Mahogany Fill Solid Border", _
      StringAlignmentCenter, White
   DrawGDIPFocusRect hDC, 178, 216, 77, 28, _
      1, White, DashStyleDot
   DrawShell32Icon hDC, 4, 144, 214
   
   '##### RR 6
   DrawGDIPMultiStyleRectangle hDC, _
      140, 260, 260, 300, 4, _
      BrushTypeTextureFill, _
      3, PenTypeSolidColor, DarkBrown, XPGoldLight, _
      XPGoldLight, XPGoldDark, , _
      AppPath & "Spruce.gif"
   DrawGDIPFormattedText hDC, 178, 266, 77, 28, _
      "Spruce Fill 3D Sunken Border", _
      StringAlignmentCenter, White
   DrawGDIPFocusRect hDC, 178, 266, 77, 28, _
      1, White, DashStyleDot
   DrawShell32Icon hDC, 4, 144, 264

End Sub

Private Sub DrawLineCaps()
   Dim graphics         As Long, pen As Long

   ' Initializations
   GdipCreateFromHDC hDC, graphics      ' Initialize the graphics class - required for all drawing
   GdipCreatePen1 Blue, 8, UnitPixel, pen

   ' Set the start and end caps
   GdipSetPenStartCap pen, LineCapArrowAnchor
   GdipSetPenEndCap pen, LineCapRoundAnchor

   ' Draw the line
   GdipDrawLineI graphics, pen, 20, 175, 300, 175

   ' Cleanup
   GdipDeletePen pen
   GdipDeleteGraphics graphics
End Sub


Private Sub DrawCustomDashed()
   Dim graphics         As Long, pen As Long
   Dim dashValues(1 To 4) As Single

   ' Set the dash intervals
   ' The dashes are in an on/off pattern and continually repeat for the length of the line
   dashValues(1) = 5    ' Show 5 * penwidth
   dashValues(2) = 2    ' Hide 2 * penwidth
   dashValues(3) = 15   ' Show 15 * penwidth
   dashValues(4) = 4    ' Hide 4 * pendwith

   ' Initializations
   GdipCreateFromHDC hDC, graphics      ' Initialize the graphics class - required for all drawing
   GdipCreatePen1 Black, 4, UnitPixel, pen

   ' Set the dash pattern
   GdipSetPenDashArray pen, dashValues(1), 4

   ' Draw the line
   GdipDrawLineI graphics, pen, 5, 5, 405, 5

   ' Cleanup
   GdipDeletePen pen
   GdipDeleteGraphics graphics
End Sub

Private Sub DrawTexturedLine()
   Dim graphics         As Long, img As Long, pen As Long, tBrush As Long
   Dim lngHeight        As Long, lngWidth As Long

   ' Initializations
   GdipCreateFromHDC hDC, graphics      ' Initialize the graphics class - required for all drawing
   GdipLoadImageFromFile StrConv(AppPath & "Texture.bmp", vbUnicode), img
   GdipCreateTexture img, WrapModeTile, tBrush  ' Create a textured brush
   GdipCreatePen2 tBrush, 30, UnitPixel, pen   ' Create a pen to draw with

   ' Get the image height and width
   GdipGetImageHeight img, lngHeight
   GdipGetImageWidth img, lngWidth

   GdipDrawImageRect graphics, img, 0, 0, lngWidth, lngHeight
   GdipDrawEllipseI graphics, pen, 100, 20, 200, 100

   ' Cleanup
   GdipDeletePen pen
   GdipDeleteBrush tBrush
   GdipDisposeImage img
   GdipDeleteGraphics graphics
End Sub

Private Sub DrawSolidShape()
   Dim graphics         As Long, brush As Long
   Dim lngHeight        As Long, lngWidth As Long

   ' Initializations
   GdipCreateFromHDC hDC, graphics  ' Initialize the graphics class - required for all drawing
   GdipCreateSolidFill DeepPink, brush      ' Create the solid colored brush
   GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
   ' Draw an ellipse
   GdipFillEllipseI graphics, brush, 0, 0, 100, 60

   ' Cleanup
   GdipDeleteBrush brush
   GdipDeleteGraphics graphics
End Sub



Private Sub blend(RGB1 As Long, RGB2 As Long, num, Colours() As Long)
   'blends two colors together by a certain percent (decimal percent)
   Dim R                As Long
   Dim R1               As Long
   Dim r2               As Long
   Dim g                As Long
   Dim G1               As Long
   Dim G2               As Long
   Dim B                As Long
   Dim B1               As Long
   Dim B2               As Long
   Dim Percent          As Single
   Dim n                As Long

   Colours(0) = RGB1
   Colours(num) = RGB2

   R1 = RGB1 And &HFF&
   G1 = RGB1 \ 256 And &HFF
   B1 = RGB1 \ 65536 And &HFF

   r2 = RGB2 And &HFF&
   G2 = RGB2 \ 256 And &HFF
   B2 = RGB2 \ 65536 And &HFF

   For n = 1 To num - 1
      Percent = n / num
      R = (r2 - R1) * Percent + R1
      g = (G2 - G1) * Percent + G1
      B = (B2 - B1) * Percent + B1
      Colours(n) = &HFF000000 + B + 256& * g + 65536 * R
   Next

End Sub

Private Sub DrawHatchShape()
   Dim graphics         As Long, brush As Long
   Dim lngHeight        As Long, lngWidth As Long

   ' Initializations
   GdipCreateFromHDC hDC, graphics  ' Initialize the graphics class - required for all drawing
   GdipCreateHatchBrush HatchStyleDottedGrid, Black, DeepPink, brush       ' Create the pattern brush

   ' Draw an ellipse
   GdipFillEllipseI graphics, brush, 0, 0, 100, 60

   ' Cleanup
   GdipDeleteBrush brush
   GdipDeleteGraphics graphics
End Sub

Private Sub DrawCurves()
   Dim graphics         As Long, pen As Long
   Dim points(1 To 5)   As POINTL

   ' Random values (From the SDK C++ sample)
   points(1).x = 0
   points(1).y = 100
   points(2).x = 50
   points(2).y = 80
   points(3).x = 100
   points(3).y = 20
   points(4).x = 150
   points(4).y = 80
   points(5).x = 200
   points(5).y = 100

   ' Initializations
   GdipCreateFromHDC hDC, graphics      ' Initialize the graphics class - required for all drawing
   GdipCreatePen1 Black, 2, UnitPixel, pen   ' Create a pen to draw with
   ' Gamma correction is nice, though slower...
   GdipSetCompositingQuality graphics, CompositingQualityGammaCorrected
   ' Draw the curve w/ anti-aliasing
   GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
   GdipDrawCurveI graphics, pen, points(1), 5

   ' Cleanup
   GdipDeletePen pen
   GdipDeleteGraphics graphics
End Sub

Private Sub DrawScaling()
   ' Why did I store the picture in a hidden PictureBox, you ask?
   ' Well, I'll tell you why: To make the demo simple yet complex
   ' enough for you to figure out how to build on it. (I hate VERY
   ' simple demos, and you should too!) I hope this isn't too simple...
   Dim graphics         As Long, img As Long
   Dim lngHeight        As Long, lngWidth As Long

   ' We are going to draw on the form, hence the hdc
   GdipCreateFromHDC hDC, graphics   ' Initialize the graphics class - required for all drawing

   ' Load the bitmap file into the Picture box (could also embed it)
   Set picGrapes.Picture = LoadPicture(AppPath & "GrapeBunch.jpg")
   ' WARNING: Make sure the picture box is large enough - WSIWYG!
   picGrapes.AutoSize = True  ' This should do for what we need
   ' Get the image "class" from the PictureBox
   GdipCreateBitmapFromHBITMAP picGrapes.image.Handle, picGrapes.image.hpal, img
   ' Below is the "cheap" way (via file); good for all supported file type
   ' Could also use GdipCreateBitmapFromFile for this bitmap
   ' Comment out the picture box code above and uncomment this to try it out if you want!
   'GdipLoadImageFromFile(StrConv(apppath & "GrapeBunch.jpg", vbUnicode), img)

   ' Get the image height and width
   GdipGetImageHeight img, lngHeight
   GdipGetImageWidth img, lngWidth

   '**** If you don't pass a width and height when drawing, the image is auto-scaled!! ****
   'GdipDrawImage graphics, img, 10, 10) ' Auto-Scaled

   ' Draw the image with no shrinking or stretching
   GdipDrawImageRectI graphics, img, 10, 10, lngWidth, lngHeight

   ' Shrink the image using low-quality interpolation.
   GdipSetInterpolationMode graphics, InterpolationModeNearestNeighbor
   GdipDrawImageRectRectI graphics, img, 10, 250, 0.6 * lngWidth, 0.6 * lngHeight, 0, 0, lngWidth, lngHeight, UnitPixel

   ' Shrink the image using medium-quality interpolation.
   GdipSetInterpolationMode graphics, InterpolationModeHighQualityBilinear
   GdipDrawImageRectRectI graphics, img, 150, 250, 0.6 * lngWidth, 0.6 * lngHeight, 0, 0, lngWidth, lngHeight, UnitPixel

   ' Shrink the image using high-quality interpolation.
   GdipSetInterpolationMode graphics, InterpolationModeHighQualityBicubic
   GdipDrawImageRectRectI graphics, img, 290, 250, 0.6 * lngWidth, 0.6 * lngHeight, 0, 0, lngWidth, lngHeight, UnitPixel

   ' NOTES: Since we are shrinking the entire image, we could just as well have called
   '        the GdipDrawImageRectI function, which would simply things - but our goal must
   '        be make life hellish!

   ' Cleanup
   GdipDisposeImage img ' Delete the image
   GdipDeleteGraphics graphics
End Sub

Private Sub DrawSkewed()
   Dim graphics         As Long, img As Long
   Dim destinationPoints(1 To 3) As POINTL
   Dim lngHeight        As Long, lngWidth As Long

   ' Set the skewing points in the point array.
   ' destination for upper-left point of original
   destinationPoints(1).x = 200
   destinationPoints(1).y = 20
   ' destination for upper-right point of original
   destinationPoints(2).x = 110
   destinationPoints(2).y = 100
   ' destination for lower-left point of original
   destinationPoints(3).x = 250
   destinationPoints(3).y = 30

   ' Initializations
   GdipCreateFromHDC hDC, graphics  ' Initialize the graphics class - required for all drawing
   GdipLoadImageFromFile StrConv(AppPath & "Stripes.bmp", vbUnicode), img    ' Load the image

   ' Get the image height and width
   GdipGetImageHeight img, lngHeight
   GdipGetImageWidth img, lngWidth

   ' Draw the image unaltered with its upper-left corner at (0, 0).
   GdipDrawImageRectI graphics, img, 0, 0, lngWidth, lngHeight

   ' Draw the image mapped to the parallelogram.
   GdipDrawImagePointsI graphics, img, destinationPoints(1), 3

   ' Cleanup
   GdipDisposeImage img ' Delete the image
   GdipDeleteGraphics graphics
End Sub

Private Sub DrawThumbnail()
   Dim graphics         As Long, img As Long, imgThumb As Long
   Dim lngHeight        As Long, lngWidth As Long

   ' Initializations
   GdipCreateFromHDC hDC, graphics  ' Initialize the graphics class - required for all drawing
   GdipLoadImageFromFile StrConv(AppPath & "SomeBigImage.bmp", vbUnicode), img  ' Load the image

   ' Get the image height and width
   GdipGetImageHeight img, lngHeight
   GdipGetImageWidth img, lngWidth

   ' Create the thumbnail that is 100x100 in size
   GdipGetImageThumbnail img, 100, 100, imgThumb

   ' Draw the thumbnail image unaltered
   GdipDrawImageRectI graphics, imgThumb, 10, 10, lngWidth, lngHeight

   ' Cleanup
   GdipDisposeImage img ' Delete the image
   GdipDisposeImage (imgThumb) ' Delete the thumbnail image
   GdipDeleteGraphics graphics
End Sub

Private Sub DrawHorizGradient()
   Dim graphics         As Long, brush As Long, pen As Long
   Dim pt1              As POINTL, pt2 As POINTL

   ' Set the gradient color points
   pt1.x = 0
   pt1.y = 10
   pt2.x = 200
   pt2.y = 10

   ' Initializations
   GdipCreateFromHDC hDC, graphics  ' Initialize the graphics class - required for all drawing
   ' Create the gradient brush; we'll use tiling
   GdipCreateLineBrushI pt1, pt2, Red, Blue, WrapModeTile, brush
   ' Create a pen with the same gradient brush
   GdipCreatePen2 brush, 1, UnitPixel, pen

   ' Draw some objects
   GdipDrawLine graphics, pen, 0, 10, 200, 10
   GdipFillEllipse graphics, brush, 0, 30, 200, 100
   GdipFillRectangle graphics, brush, 0, 155, 500, 30

   'Cleanup
   GdipDeletePen pen
   GdipDeleteBrush brush
   GdipDeleteGraphics graphics
End Sub

Private Sub DrawDiagGradient()
   Dim graphics         As Long, brush As Long, pen As Long
   Dim pt1              As POINTL, pt2 As POINTL

   ' Set the gradient color points
   ' pt1 will stay at 0,0
   pt2.x = 200
   pt2.y = 100

   ' Initializations
   GdipCreateFromHDC hDC, graphics  ' Initialize the graphics class - required for all drawing
   ' Create the gradient brush; we'll use tiling
   GdipCreateLineBrushI pt1, pt2, Cyan, Blue, WrapModeTile, brush
   ' Create a pen with the same gradient brush
   GdipCreatePen2 brush, 10, UnitPixel, pen

   ' Draw some objects
   GdipDrawLineI graphics, pen, 0, 0, 600, 300
   GdipFillEllipseI graphics, brush, 10, 100, 200, 100

   'Cleanup
   GdipDeletePen pen
   GdipDeleteBrush brush
   GdipDeleteGraphics graphics
End Sub

Private Sub DrawPathGradient()
   Dim graphics         As Long
   Dim brush            As Long
   Dim path             As Long

   ' Initializations
   GdipCreateFromHDC hDC, graphics  ' Initialize the graphics class - required for all drawing
   'AA
   GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
   ' Create a GraphicsPath object
   GdipCreatePath FillModeWinding, path

   ' Add an ellipse to the path
   GdipAddPathEllipseI path, 0, 0, 140, 70

   ' Create a path gradient based on the ellipse
   GdipCreatePathGradientFromPath path, brush

   ' Set the middle color of the path to Blue
   GdipSetPathGradientCenterColor brush, Blue

   ' Set the entire path boundary to Aqua
   ' NOTE: This expects an array, but since we only have one item we can fudge it
   GdipSetPathGradientSurroundColorsWithCount brush, Aqua, 1

   ' Draw the ellipse, keeping the exact coords we defined for the path
   GdipFillEllipse graphics, brush, 0, 0, 140, 70

   'Cleanup
   GdipDeletePath path    ' Delete the path object
   GdipDeleteBrush brush
   GdipDeleteGraphics graphics
End Sub

' NOTE: Use this same concept for: BMP, GIF, and PNG format saving
Private Sub BMPtoPNG()
   Dim img              As Long, encoderCLSID As CLSID
   Dim stat             As GpStatus

   ' Initializations
   ' No graphics object needed here since we aren't doing any drawing.
   ' We'll convert the grapes bitmap file
   GdipLoadImageFromFile StrConv(AppPath & "GrapeBunch.jpg", vbUnicode), img

   ' Get the CLSID of the PNG encoder
   GetEncoderClsid "image/png", encoderCLSID

   ' Save as a PNG file. There are no encoder parameters for PNG images, so we pass a NULL.
   ' NOTE: The NULL (aka 0) must be passed byval, as the function declaration would get a pointer to the number 0.
   stat = GdipSaveImageToFile(img, StrConv(AppPath & "GrapeBunch.png", vbUnicode), encoderCLSID, ByVal 0)

   ' See if it was created
   If stat = Ok Then
      MsgBox "Successfully saved GrapeBunch.png!", vbInformation
   Else
      MsgBox "Error saving file! Status Code: " & stat, vbCritical
   End If

   ' Cleanup
   GdipDisposeImage img
End Sub

' Note: Use this same concept for: JPEG and TIFF saving
Private Sub BMPtoJPEG()
   Dim img              As Long, encoderCLSID As CLSID
   Dim stat             As GpStatus
   Dim encoderParams    As EncoderParameters

   ' Initializations
   ' No graphics object needed here since we aren't doing any drawing.
   ' We'll convert the grapes bitmap file
   GdipLoadImageFromFile StrConv(AppPath & "GrapeBunch.jpg", vbUnicode), img

   ' Get the CLSID of the PNG encoder
   GetEncoderClsid "image/jpeg", encoderCLSID

   ' Save as a JPEG file. This format requires encoder parameters.
   ' Setup the encoder paramters
   encoderParams.count = 1    ' Only one element in this Parameter array
   With encoderParams.Parameter
      .NumberOfValues = 1     ' Should be one
      .type = EncoderParameterValueTypeLong
      ' Set the GUI to EncoderQuality
      GetParameterGUID EncoderQuality, .GUID
      .value = VarPtr(90)  ' Remember: The value expects only pointers!
   End With

   ' Now save the bitmap as a jpeg at 10% compression
   stat = GdipSaveImageToFile(img, StrConv(AppPath & "GrapeBunch.jpg", vbUnicode), encoderCLSID, encoderParams)

   ' See if it was created
   If stat = Ok Then
      MsgBox "Successfully saved GrapeBunch.jpg!", vbInformation
   Else
      MsgBox "Error saving file! Status Code: " & stat, vbCritical
   End If

   ' Cleanup
   GdipDisposeImage img
End Sub

' Now that we know how to set the value of one encoding parameter, what do we do if we
' want to set more than one encoding parameter? Well, this function will show you how to
' do it!
' Note: Requires the CopyMemory API
' >> NOTE: You can ONLY rotate JPEG images! If you load a NON-JPEG image and try to rotate, nothing will happen!
Private Sub BMPtoJPEG_Params()
   Dim img              As Long, encoderCLSID As CLSID
   Dim stat             As GpStatus
   Dim encoderParams    As EncoderParameters ' This will now become a temporary holder
   Dim encoderArray()   As Byte             ' Our main "struct"
   Dim lngEP            As Long                        ' Size of encoderParams variable/struct

   ' Initializations
   ' No graphics object needed here since we aren't doing any drawing.
   ' We'll rotate the GrapeBunch.jpg file
   GdipLoadImageFromFile StrConv(AppPath & "GrapeBunch.jpg", vbUnicode), img
   lngEP = Len(encoderParams)

   ' Get the CLSID of the PNG encoder
   GetEncoderClsid "image/jpeg", encoderCLSID

   ' Determine how many parameters we will need
   ' JPEGs can only use 2 parameters, so we'll use both
   ReDim encoderArray(0 To (lngEP + Len(encoderParams.Parameter))) As Byte

   ' Save as a JPEG file. This format requires encoder parameters.
   ' Setup the encoder paramters
   ' We'll setup the struct and first parameter as usual
   encoderParams.count = 2    ' We are setting 2 parameters
   With encoderParams.Parameter
      .NumberOfValues = 1     ' Should be one
      .type = EncoderParameterValueTypeLong
      ' Set the GUI to EncoderQuality
      GetParameterGUID EncoderQuality, .GUID
      .value = VarPtr(100)  ' Remember: The value expects only pointers!
   End With

   ' Copy the data into the byte array
   CopyMemory encoderArray(0), encoderParams, lngEP

   ' Now we'll re-use the parameter member of encoderParams
   With encoderParams.Parameter
      .NumberOfValues = 1     ' Should be one
      .type = EncoderParameterValueTypeLong
      ' Set the GUI to EncoderTransformation
      GetParameterGUID EncoderTransformation, .GUID
      ' We'll flip horizontally - REMEMBER TO USE A POINTER!
      .value = VarPtr(EncoderValueTransformRotate180)
   End With

   ' Copy the second parameter to the byte array at the right offset
   CopyMemory encoderArray(lngEP), encoderParams.Parameter, Len(encoderParams.Parameter)

   ' Now save the bitmap as a jpeg at 0% compression to try to keep the quality up
   ' Notice how the byte array is passed instead of the struct
   stat = GdipSaveImageToFile(img, StrConv(AppPath & "GrapeBunch180.jpg", vbUnicode), encoderCLSID, encoderArray(0))

   ' See if it was created
   If stat = Ok Then
      MsgBox "Successfully saved GrapeBunch.jpg!", vbInformation
   Else
      MsgBox "Error saving file! Status Code: " & stat, vbCritical
   End If

   ' Cleanup
   GdipDisposeImage img
End Sub

Private Sub DrawCachedBitmap()
   Dim graphics         As Long, bitmap As Long, cBitmap As Long
   Dim lngHeight        As Long, lngWidth As Long
   Dim j                As Long, k As Long

   ' Initializations
   GdipCreateFromHDC hDC, graphics  ' Initialize the graphics class - required for all drawing
   GdipLoadImageFromFile StrConv("Texture.bmp", vbUnicode), bitmap      ' Load the image
   ' Create a cached bitmap from the loaded image
   GdipCreateCachedBitmap bitmap, graphics, cBitmap

   ' Get the image height and width
   GdipGetImageHeight bitmap, lngHeight
   GdipGetImageWidth bitmap, lngWidth

   ' Perform a test to see which is faster
   For j = 1 To 300 Step 10
      For k = 1 To 1000
         GdipDrawImageRect graphics, bitmap, j, j / 2, lngWidth, lngHeight
      Next
   Next

   For j = 1 To 300 Step 10
      For k = 1 To 1000
         GdipDrawCachedBitmap graphics, cBitmap, j, 150 + j / 2
      Next
   Next

   ' Cleanup
   GdipDisposeImage bitmap
   GdipDeleteCachedBitmap cBitmap  ' Note the special deletion function
   GdipDisposeImage cBitmap        ' This may not be needed...
   GdipDeleteGraphics graphics
End Sub

Private Sub DrawAlphaLines()
   Dim graphics         As Long
   Dim bitmap           As Long
   Dim lngHeight        As Long
   Dim lngWidth         As Long
   Dim opaquePen        As Long
   Dim semiTansPen      As Long

   ' Initializations
   GdipCreateFromHDC hDC, graphics  ' Initialize the graphics class - required for all drawing
   GdipLoadImageFromFile StrConv(AppPath & "Texture.bmp", vbUnicode), bitmap   ' Load the image
   GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
   ' Create our pens for line drawing
   GdipCreatePen1 ColorARGB(255, 0, 0, 255), 2, UnitPixel, opaquePen
   GdipCreatePen1 ColorARGB(128, 0, 0, 255), 15, UnitPixel, semiTansPen  ' Has 50% alpha blending

   ' Get the image height and width
   GdipGetImageHeight bitmap, lngHeight
   GdipGetImageWidth bitmap, lngWidth

   ' Draw the image without auto-scaling
   GdipDrawImageRect graphics, bitmap, 10, 5, lngWidth, lngHeight

   ' Draw an opaque line over the image
   GdipDrawLine graphics, opaquePen, 0, 20, 100, 40
   ' Draw the semi-transparent line over the image
   GdipDrawLine graphics, semiTansPen, 0, 40, 100, 80
   ' Draw the same semi-transparent line, but with gamma correction
   GdipSetCompositingQuality graphics, CompositingQualityGammaCorrected
   GdipDrawLine graphics, semiTansPen, 0, 60, 100, 120

   ' Cleanup
   GdipDeletePen opaquePen
   GdipDeletePen semiTansPen
   GdipDisposeImage bitmap
   GdipDeleteGraphics graphics
End Sub

Private Sub DrawColorMatrix()
   Dim graphics         As Long, bitmap As Long, pen As Long
   Dim imgAttr          As Long, clrMatrix As ColorMatrix
   Dim lngHeight        As Long, lngWidth As Long

   ' Initializations
   GdipCreateFromHDC hDC, graphics  ' Initialize the graphics class - required for all drawing
   GdipLoadImageFromFile StrConv(AppPath & "Texture.bmp", vbUnicode), bitmap    ' Load the image
   GdipCreatePen1 Black, 15, UnitPixel, pen   ' Create an opaque pen

   ' Get the image height and width
   GdipGetImageHeight bitmap, lngHeight
   GdipGetImageWidth bitmap, lngWidth

   ' Fill the color matrix
   ' Notice the value 0.8 in row 4, column 4.
   clrMatrix.m(0, 0) = 1: clrMatrix.m(1, 0) = 0: clrMatrix.m(2, 0) = 0: clrMatrix.m(3, 0) = 0: clrMatrix.m(4, 0) = 0
   clrMatrix.m(0, 1) = 0: clrMatrix.m(1, 1) = 1: clrMatrix.m(2, 1) = 0: clrMatrix.m(3, 1) = 0: clrMatrix.m(4, 1) = 0
   clrMatrix.m(0, 2) = 0: clrMatrix.m(1, 2) = 0: clrMatrix.m(2, 2) = 1: clrMatrix.m(3, 2) = 0: clrMatrix.m(4, 2) = 0
   clrMatrix.m(0, 3) = 0: clrMatrix.m(1, 3) = 0: clrMatrix.m(2, 3) = 0: clrMatrix.m(3, 3) = 0.8: clrMatrix.m(4, 3) = 0
   clrMatrix.m(0, 4) = 0: clrMatrix.m(1, 4) = 0: clrMatrix.m(2, 4) = 0: clrMatrix.m(3, 4) = 0: clrMatrix.m(4, 4) = 1

   ' Create the ImageAttributes object
   GdipCreateImageAttributes imgAttr
   ' And set its color matrix
   GdipSetImageAttributesColorMatrix imgAttr, ColorAdjustTypeDefault, True, clrMatrix, ByVal 0, ColorMatrixFlagsDefault

   ' Draw a wide black line
   GdipDrawLine graphics, pen, 10, 35, 200, 35

   ' Draw the semi-transparent image
   GdipDrawImageRectRectI graphics, bitmap, 30, 0, lngWidth, lngHeight, 0, 0, lngWidth, lngHeight, UnitPixel, imgAttr

   ' Cleanup
   GdipDisposeImageAttributes imgAttr ' Delete the Image attributes object
   GdipDeletePen pen
   GdipDisposeImage bitmap
   GdipDeleteGraphics graphics
End Sub

' The slower way of using alpha-blending
Private Sub DrawAlphaPixels()
   Dim graphics         As Long, bitmap As Long, pen As Long
   Dim lngHeight        As Long, lngWidth As Long
   Dim iRow             As Long, iCol As Long, lARGB As Long

   ' Initializations
   GdipCreateFromHDC hDC, graphics  ' Initialize the graphics class - required for all drawing
   GdipLoadImageFromFile StrConv(AppPath & "Texture.bmp", vbUnicode), bitmap  ' Load the image
   GdipCreatePen1 Black, 15, UnitPixel, pen   ' Create an opaque pen

   ' Get the image height and width
   GdipGetImageHeight bitmap, lngHeight
   GdipGetImageWidth bitmap, lngWidth

   ' Modify the pixels in the bitmap
   ' NOTE: I'm pretty sure that the bitmap object it forever modified by doing this.
   '       If you still want the original, I would suggest cloning this image first.
   For iRow = 0 To (lngHeight - 1)
      For iCol = 0 To (lngWidth - 1)
         ' Get the current ARGB color of the pixel
         GdipBitmapGetPixel bitmap, iCol, iRow, lARGB
         ' Set the pixel color back with a new alpha
         ' NOTE: I created a helper function for alpha setting to make it easier
         GdipBitmapSetPixel bitmap, iCol, iRow, ColorSetAlpha(lARGB, 255 * iCol / lngWidth)
      Next
   Next

   ' Draw a wide black line
   GdipDrawLine graphics, pen, 10, 35, 200, 35

   ' Draw the modified image
   GdipDrawImageRect graphics, bitmap, 30, 0, lngWidth, lngHeight

   ' Cleanup
   GdipDeletePen pen
   GdipDisposeImage bitmap
   GdipDeleteGraphics graphics
End Sub

Private Sub DrawAntiAliasText()
   Dim graphics         As Long, brush As Long
   Dim fontFam          As Long, curFont As Long
   Dim rcLayout         As RECTF   ' Designates the string drawing bounds

   ' Initializations
   GdipCreateFromHDC hDC, graphics  ' Initialize the graphics class - required for all drawing
   GdipCreateSolidFill Blue, brush     ' Create a brush to draw the text with
   ' Create a font family object to allow use to create a font
   ' We have no font collection here, so pass a NULL for that parameter
   GdipCreateFontFamilyFromName StrConv("Times New Roman", vbUnicode), 0, fontFam
   ' Create the font from the specified font family name
   GdipCreateFont fontFam, 32, FontStyleRegular, UnitPixel, curFont

   ' Set up a drawing area
   ' NOTE: Leaving the right and bottom values at zero means there is no boundary
   rcLayout.Left = 10
   rcLayout.Top = 10

   ' This function allows us to alter the text quality.
   ' We'll use the worst quality first.
   GdipSetTextRenderingHint graphics, TextRenderingHintSingleBitPerPixel
   ' We have no string format object, so pass a NULL for that parameter
   GdipDrawString graphics, StrConv("SingleBitPerPixel", vbUnicode), 17, curFont, rcLayout, 0, brush

   ' Set up another drawing area
   rcLayout.Left = 10
   rcLayout.Top = 60

   ' Now we'll use anti-aliasing
   GdipSetTextRenderingHint graphics, TextRenderingHintAntiAlias
   ' We have no string format object, so pass a NULL for that parameter
   GdipDrawString graphics, StrConv("AntiAlias", vbUnicode), 9, curFont, rcLayout, 0, brush

   ' Cleanup
   GdipDeleteFont (curFont)    ' Delete the font object
   GdipDeleteFontFamily (fontFam) ' Delete the font family object
   GdipDeleteBrush brush
   GdipDeleteGraphics graphics
End Sub

' Note: This example was inspired by another post on planetsourcecode today.
'       Someone asked if GDI+ could rotate images with anti-alias, and behold!
Private Sub DrawRotated()
   Dim graphics         As Long
   Dim img              As Long
   Dim pen              As Long
   Dim lngHeight        As Long
   Dim lngWidth         As Long

   ' Initializations
   GdipCreateFromHDC hDC, graphics  ' Initialize the graphics class - required for all drawing
   GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
   GdipLoadImageFromFile StrConv(AppPath & "GrapeBunch.jpg", vbUnicode), img

   ' Get the image height and width
   GdipGetImageHeight img, lngHeight
   GdipGetImageWidth img, lngWidth

   ' This will rotate EVERYTHING!
   ' There are several rotation APIs available for you!
   GdipRotateWorldTransform graphics, 65, MatrixOrderAppend

   ' Make sure to provide a good x,y starting point!
   GdipDrawImageRect graphics, img, 200, -150, lngWidth, lngHeight

   ' Cleanup
   GdipDisposeImage img
   GdipDeleteGraphics graphics
End Sub

Private Sub Timer1_Timer()
   
   Dim I As Long
   
   lAniStep = (lAniStep + 1) Mod 7

   For I = 0 To 15
      Stars(I).x = Stars(I).x + 2
      If Stars(I).x > 119 Then
         Stars(I).x = 1
      End If
      Stars(I).y = Stars(I).y + 1
      If Stars(I).y > 39 Then
         Stars(I).y = 1
      End If
   Next

   'This is just a demo. If this were an ActiveX Control
   'then you would just refresh the button here and use a
   'callback event to paint the background or animation
   
   DrawGDIPMultiStyleRectangle hDC, _
      140, 10, 260, 52, 4, _
      BrushTypeLinearGradient, _
      1, PenTypeSolidColor, XPBlue, , _
      White, XPGradient
   DrawGDIPFormattedText hDC, 179, 14, 77, 32, _
      "XP Style GDI+ Animated", _
      StringAlignmentCenter, Black
   DrawGDIPFocusRect hDC, 144, 14, 112, 34, _
      1, Black, DashStyleDot
   DrawIconEx hDC, 146, 16, img(lAniStep).Picture.Handle, _
         32, 32, 0, 0, DI_NORMAL
         
'#####

   DrawGDIPMultiStyleRectangle hDC, _
      10, 60, 130, 100, 0, _
      BrushTypeLinearGradient, _
      1, PenTypeSolidColor, White, , _
      DarkBlue, Blue
   
   For I = 0 To 15
      SetPixelV hDC, Stars(I).x + 10, Stars(I).y + 60, vbWhite
   Next
   
   DrawShell32Icon hDC, 20, 14, 64
   
   DrawGDIPFormattedText hDC, 48, 62, 77, 36, _
      "Owner Draw Background", _
      StringAlignmentCenter, White
         
   Me.Refresh
   
End Sub
