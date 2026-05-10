Option Explicit

' ==============================================================================
' EXCEL IMAGE INSERTION TOOL
' Instructions: Paste this entire code into the "ThisWorkbook" module.
' ==============================================================================

Public IsImageMode As Boolean
Private LastClickTime As Double

'''
''' Logic to Toggle Mode using 4-Clicks (Two rapid Double-Clicks)
'''
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim CurrentTime As Double
    CurrentTime = Timer
    
    ' Handle Timer reset at midnight (CurrentTime becomes smaller than LastClickTime)
    Dim TimeDiff As Double
    If CurrentTime < LastClickTime Then
        TimeDiff = (86400 - LastClickTime) + CurrentTime
    Else
        TimeDiff = CurrentTime - LastClickTime
    End If
    
    ' If the time between two Double-Clicks is less than 1 second, toggle mode
    If TimeDiff < 1 And LastClickTime > 0 Then
        Cancel = True ' Stop cell editing
        IsImageMode = Not IsImageMode ' Toggle Mode
        
        If IsImageMode Then
            MsgBox "Image Insertion Mode: [ON]" & vbCrLf & _
                   "------------------------------------" & vbCrLf & _
                   "- Right-click any cell to insert an image." & vbCrLf & _
                   "- Images will auto-fit to the cell size.", vbInformation, "System - Image Mode"
        Else
            MsgBox "Image Insertion Mode: [OFF]" & vbCrLf & _
                   "------------------------------------" & vbCrLf & _
                   "- Right-click menu returned to normal.", vbExclamation, "System - Image Mode"
        End If
        
        LastClickTime = 0 ' Reset timer
    Else
        ' First Double-click detected or too slow, start timer
        LastClickTime = CurrentTime
        
        ' If mode is ON, we prevent editing even on first double-click for consistency
        If IsImageMode Then Cancel = True
    End If
End Sub

'''
''' Right-click to insert Image (Supports .webp, .jpeg, .png, etc.)
'''
Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If IsImageMode Then
        Cancel = True ' Suppress default context menu
        
        Dim strFilePath As Variant
        Dim img As Shape
        Dim targetAddr As String
        targetAddr = Target.Address
        
        ' Open File Dialog with expanded filters
        strFilePath = Application.GetOpenFilename( _
            FileFilter:="Image Files (*.webp;*.jpg;*.jpeg;*.png;*.bmp),*.webp;*.jpg;*.jpeg;*.png;*.bmp", _
            Title:="Select an Image to Insert into " & targetAddr)
        
        ' If user didn't cancel
        If strFilePath <> False Then
            On Error Resume Next ' Simple error handling for file access issues
            
            ' 1. Clean up existing images exactly in the target cell
            For Each img In Sh.Shapes
                If img.TopLeftCell.Address = targetAddr Then
                    img.Delete
                End If
            Next img

            ' 2. Insert the picture
            ' Note: Using Width:=-1 and Height:=-1 to load original dimensions first
            Set img = Sh.Shapes.AddPicture(Filename:=strFilePath, _
                                           LinkToFile:=msoFalse, _
                                           SaveWithDocument:=msoCTrue, _
                                           Left:=Target.Left, _
                                           Top:=Target.Top, _
                                           Width:=-1, _
                                           Height:=-1)
            
            If Err.Number <> 0 Then
                MsgBox "Error inserting image: " & Err.Description, vbCritical, "Error"
                Err.Clear
                On Error GoTo 0
                Exit Sub
            End If
            
            ' 3. Fit to cell with a 1-pixel padding
            With img
                .LockAspectRatio = msoFalse ' Stretch to fit cell
                .Left = Target.Left + 1
                .Top = Target.Top + 1
                .Width = Target.Width - 2
                .Height = Target.Height - 2
                .Placement = xlMoveAndSize ' Ensure it moves/sizes with cells
            End With
            
            On Error GoTo 0
        End If
    End If
End Sub
