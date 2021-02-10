# VBA-Projects

This is a VBA macro I wrote to address an issue with how enrollment reports displayed one course that had multiple cross-listings.  In order for one department to know how many students were enrolled in different courses, one would have to find all the cross-listed sections and then manually add them together.  This macro would automate that process across hundreds of cross-listed courses among all disciplines and campuses.

```VBA
Option Explicit

Sub CombineCrosslistedCourses()

    Dim nRows, i, j, cell1, cell2, MaxEnroll, CurrentEnroll As Variant

'Sort Courses by CrossListing
    Range("A1").Select
    
    Range("A1").CurrentRegion.Sort Key1:=Range("F1"), Order1:=xlAscending, Header:=xlYes

'Combine enrollment by CrossListing


    'Select Data Range
    Selection.CurrentRegion.Select
    
    nRows = Selection.Rows.Count
    
    'create copy of SSRXLST_XLST_GROUP (crosslisting code)
    Columns("F:F").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("Y:Y").Select
    ActiveSheet.Paste
    
    
    'convert copy of SSRXLST_XLST_GROUP (crosslisting code) to text so that non-crosslisted codes are not combined
    Columns("Y:Y").Select
    Selection.NumberFormat = "@" ‘convert Xlist codes to text - $ gets lost as it is read as symbol and not text.
    
    'loop thru rows
    
    For i = 2 To nRows
    
        'save crosslisted code in cell1
        
        cell1 = Cells(i, 25)
            
        'check that cell1 is crosslisted
        
        If (cell1 <> "") Then
            'loop thru data range to compare other courses to course saved in cell1
            For j = 2 To nRows
                'save current course crosslist code in cell2
                cell2 = Cells(j, 25)
                
                'check that cell2 is crosslisted
                If (cell2 <> "") Then
                    
                    'check crosslist codes and subject course codes match
                    If cell1 = cell2 And Cells(i, 8) = Cells(j, 8) And Cells(i, 9) = Cells(j, 9) Then
                    
                    'check if comparing course to itself
                    If i <> j Then
                        
                        MaxEnroll = Cells(j, 17)   'seat capacity of other crosslisted course
                        CurrentEnroll = Cells(j, 18)  'current enrollment of other crosslisted course
                        Cells(i, 17) = Cells(i, 17) + MaxEnroll  'combine seat capacities
                        Cells(i, 18) = Cells(i, 18) + CurrentEnroll   'combine current enrollment
                        Rows(j).Delete   'delete other crosslisted course
                        j = j - 1   'reduce current loop counter by 1 since a course was deleted
                        nRows = nRows - 1  'reduce row counter by 1 since a course was deleted
                        
                     End If  'combining course enrollment data
                  End If 'check course codes and subjects match
                End If 'current cell is crosslisted
            Next j  'pick next course to check crosslisted
        End If 'is target course crosslisted
    Next i 'pick next target course to combine



'Add Campus Columns

    nRows = Selection.CurrentRegion.Rows.Count
    
    Cells(1, 23) = "CAMPUS"
    
    For i = 2 To nRows
        ‘Read leading building code letter for assigned course classroom to determine campus
        If Left(Cells(i, 15), 1) = "C" Then Cells(i, 23) = "Clarkston"
        If Left(Cells(i, 15), 1) = "1" Then Cells(i, 23) = "Newton"
        If Left(Cells(i, 15), 1) = "2" Then Cells(i, 23) = "Newton"
        If Left(Cells(i, 15), 1) = "S" Then Cells(i, 23) = "Decatur"
        If Left(Cells(i, 15), 1) = "N" Then Cells(i, 23) = "Dunwoody"
        If Left(Cells(i, 15), 1) = "A" Then Cells(i, 23) = "Alpharetta"
        If Left(Cells(i, 15), 1) = "O" Then Cells(i, 23) = "Online"
        
    Next i

'Remove copy of SSRXLST_XLST_GROUP (crosslisting code)
    
    Columns("Y:Y").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft

End Sub
```
