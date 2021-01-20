## Resize Image and Add Hyperlinks in Word Document with VBA

This is a macro that will resize a selected image and attach a hyperlink to the image.  It will also add the hyperlink as text below the image.

To help students in my classes, I would provide YouTube links for problems similar to their assignments. 
I would create a Word doc that would have thumbnail images and links to the videos.  

The process required:
    Taking a screen snip of the video  
    Pasting it into the word document  
    Copy the url link to the clipboard.
    Manually resizing the image to 200 x 125
    Learning the "add hyperlink shortcut (Ctrl-K)"
    Inserting the hyperlink text below the image.  

Not the worse set of steps, but enough to inspire me to write this macro.

Before running the macro, select the image to be resized in the Word document and have the URL link in the Clipboard.

```VBA
Option Explicit

Sub ResizeAndLink()

'#################################################
'###  Rabbit hole warning: DataObject requires a library from Tools > References > Microsoft Forms 2.0 Object Library
'###   -If this library option is not list, your will need to download or locate the FM20.DLL file. Suggested to install in \System32
'###   -Once you download FM20.DLL open the Tools>Reference in VBA, select "Browse" and navigate to FM20.DLL
'###   -You may need to give yourself admin access to the \system32 folder, which is a whole new rabbit hole
'#################################################
  
        'Get URL from clipboard
    Dim DataObj As New MSForms.DataObject
    DataObj.GetFromClipboard
    URL = DataObj.GetText

'###########################################
'### Format Image and Add hyperlink to Image
'###########################################

        'Format selected image
With Selection.InlineShapes(1)
    .LockAspectRatio = msoFalse
    .Width = 200
    .Height = 125
        
            'other image formating options
    '.Fill.Visible = msoFalse
    '.Fill.Solid
    '.Fill.Transparency = 0#
    '.Line.Weight = 0.75
    '.Line.Transparency = 0#
    '.Line.Visible = msoFalse
    '.PictureFormat.Brightness = 0.5
    '.PictureFormat.Contrast = 0.5
    '.PictureFormat.ColorType = msoPictureAutomatic
    '.PictureFormat.CropLeft = 0#
    '.PictureFormat.CropRight = 0#
    '.PictureFormat.CropTop = 0#
    '.PictureFormat.CropBottom = 0#
    '.LockAspectRatio = msoTrue
    'Selection.InlineShapes(1).Left = 0#
    'Selection.InlineShapes(1).Top = 0#
    
         'Add hyperlink to selected image
     Selection.Hyperlinks.Add Anchor:=Selection.Range, Address:=URL, ScreenTip:="YouTube Link", TextToDisplay:=URL
       
        'Deselect image and move cursor to next line
    Selection.EndKey Unit:=wdLine
    Selection.TypeParagraph

'#################################################################################################################################
'###    Create text hyperlink
'###      -paste text from clipboard, select it, format font, add hyperlink and move to next line
'#################################################################################################################################
   
        'Paste URL text from Clipboard into document
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    
        'select URL text and copy
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    Selection.Copy
    
        'format text
    Selection.Font.Size = 8
    
        'Add hyperlink to selected URL text
    Selection.Hyperlinks.Add Anchor:=Selection.Range, Address:=URL, ScreenTip:="YouTube Link", TextToDisplay:=URL
    
        'move cursor to next line
    Selection.TypeParagraph
    
End With
End Sub
