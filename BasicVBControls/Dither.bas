Attribute VB_Name = "Dither"
Option Explicit
' Dither.bas is not my code.
' This code was downloaded from the
' Internet.  I only made changes to DitherBlue to make it DitherRed
' The original subroutine was Dither.
' To use this bas, add it to your project files
' Then in the Sub Form_Activate()
' Add the line "DitherRed Me" to make your form gradiant red
' Add the line "DitherBlue Me" to make it gradiant blue
' Set the form's "Auto Redraw" to "True"
' You can also call the subs from a raise event

Sub DitherRed(vForm As Form)
   Dim intLoop As Integer
      vForm.DrawStyle = vbInsideSolid
      vForm.DrawMode = vbCopyPen
      vForm.ScaleMode = vbPixels
      vForm.DrawWidth = 2
      vForm.ScaleHeight = 256
      For intLoop = 0 To 255
         vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
        
      Next intLoop
   End Sub
   
   
   Sub DitherBlue(vForm As Form)
   Dim intLoop As Integer
      vForm.DrawStyle = vbInsideSolid
      vForm.DrawMode = vbCopyPen
      vForm.ScaleMode = vbPixels
      vForm.DrawWidth = 2
      vForm.ScaleHeight = 256
      For intLoop = 0 To 255
         vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
      Next intLoop
   End Sub
