Attribute VB_Name = "modCommentsToExcel"
Option Explicit
Option Compare Text
Const intHeaderRow = 1

Public Sub subDeleteResolved()
    '====================================================================
    'DeleteResolved Subroutine
    '--------------------------------------------------------------------
    'Purpose    :   Delete all resolved comments
    '
    'Author     :   Callum Reid
    '
    'Notes      :
    '
    '--------------------------------------------------------------------
    'Revision History
    '--------------------------------------------------------------------
    '
    'Version 1.0.0  12/04/18    CR      - Reformatted Release
    '
    '--------------------------------------------------------------------

    'Turn off screen updating
    Application.ScreenUpdating = False
    
    'Declare variables
    Dim comment As comment
    
    For Each comment In ActiveDocument.Comments
        If comment.Done Then
            comment.DeleteRecursively
            'MsgBox comment.range
        End If
    Next comment
    
    'Turn off screen updating
    Application.ScreenUpdating = True
End Sub

Public Sub subCommentFields()
    '====================================================================
    'CommentFields Subroutine
    '--------------------------------------------------------------------
    'Purpose    :   Apply comment to all Fields that contain the string strFind
    '
    'Author     :   Callum Reid
    '
    'Notes      :
    '
    '--------------------------------------------------------------------
    'Revision History
    '--------------------------------------------------------------------
    '
    'Version 1.0.0  12/04/18    CR      - Reformatted Release
    '
    '--------------------------------------------------------------------
    
    'Turn off screen updating
    Application.ScreenUpdating = False
    
    'Declare variables
    Dim field As field
    Dim strFind As String
    Dim strComment As String
    
    'Initialise variables
    strFind = "[complete/delete]"
    
    For Each field In ActiveDocument.Fields
        If InStr(field.Code.Text, strFind) > 0 Then
            field.Select
            Selection.Collapse wdCollapseStart
            Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
            strComment = Selection.Text
            
            Selection.Collapse wdCollapseStart
            Selection.EndKey Unit:=wdLine, Extend:=wdExtend
            
            ActiveDocument.Comments.Add range:=Selection.range, Text:=strComment
        End If
    Next field
    
    'Turn off screen updating
    Application.ScreenUpdating = True
End Sub

Public Sub subExportComments()
    '====================================================================
    'ExportComments Subroutine
    '--------------------------------------------------------------------
    'Purpose    :   Export comments from Word file to Excel
    '
    'Author     :   Alexander Baekelandt
    '               Jonathan Bishop
    '               Callum Reid
    '
    'Notes      :
    '
    '--------------------------------------------------------------------
    'Revision History
    '--------------------------------------------------------------------
    '
    'Version 1.0.0  24/03/18    AB/JB   - Initial Release
    'Version 1.1.0  11/04/18    CR      - Reformatted Release
    '
    '--------------------------------------------------------------------
    
    'Turn off screen updating
    Application.ScreenUpdating = False
    
    ' Declare variables
    Dim doc As Word.Document
    Dim arrComments() As String
    Dim intComment As Integer
    Dim comment As comment
    
    ' Initialise variables
    Set doc = ActiveDocument
    
    If doc.Comments.COUNT > 0 Then
        ReDim arrComments(1 To doc.Comments.COUNT, 1 To 5)
        
        ' Iterate through all comments in active document
        For intComment = 1 To doc.Comments.COUNT
            Set comment = doc.Comments(intComment)
            arrComments(intComment, 1) = fcnGetPreviousHeadings(comment.Scope)
            arrComments(intComment, 2) = comment.Scope
            arrComments(intComment, 3) = comment.range
            arrComments(intComment, 4) = comment.Done
            arrComments(intComment, 5) = comment.Author
        Next intComment
        
        ' Export arrComments to Excel
        Call subExcelExport(arrComments)
    End If
    
    'Turn on screen updating
    Application.ScreenUpdating = True
End Sub

Private Function fcnGetPreviousHeadings(rngComment As range)
    '====================================================================
    'GetPreviousHeadings Function
    '--------------------------------------------------------------------
    'Purpose    :   Get heading that precedes given range
    '
    'Author     :   Alexander Baekelandt
    '               Jonathan Bishop
    '               Callum Reid
    '
    'Notes      :
    '
    '--------------------------------------------------------------------
    'Parameters
    '--------------------------------------------------------------------
    '
    'rngComment        : Commented text range   Range
    '--------------------------------------------------------------------
    'Revision History
    '--------------------------------------------------------------------
    '
    'Version 1.0.0  24/03/18    AB/JB   - Initial Release
    'Version 1.1.0  11/04/18    CR      - Reformatted Release
    '
    '--------------------------------------------------------------------

    ' Declare variables
    Dim rngSelection As Selection
    
    ' Go to previous heading
    rngComment.Select
    Selection.HomeKey Unit:=wdLine
    If Not InStr(Selection.Style, "Heading") <> 0 Then
        Selection.GoTo What:=wdGoToHeading, Which:=wdGoToPrevious
    End If
    Selection.Expand wdParagraph
    Set rngSelection = Selection

    ' Check that the selection is indeed a heading
    If Not InStr(1, rngSelection.Style, "Heading") <> 0 Then
        fcnGetPreviousHeadings = "No heading precedes commented text"
    Else
        fcnGetPreviousHeadings = rngSelection.Paragraphs(1).range.ListFormat.ListString & " - " & rngSelection.Text
    End If
End Function

Private Sub subExcelExport(ByRef arrComments() As String)
    '====================================================================
    'ExcelExport Subroutine
    '--------------------------------------------------------------------
    'Purpose    :   Export comments to Excel and save file
    '
    'Author     :   Alexander Baekelandt
    '               Jonathan Bishop
    '               Callum Reid
    '
    'Notes      :
    '
    '--------------------------------------------------------------------
    'Parameters
    '--------------------------------------------------------------------
    '
    'arrComments        : All comments  String Array
    '--------------------------------------------------------------------
    'Revision History
    '--------------------------------------------------------------------
    '
    'Version 1.0.0  24/03/18    AB/JB   - Initial Release
    'Version 1.1.0  11/04/18    CR      - Reformatted Release
    '
    '--------------------------------------------------------------------

    'Declare variables
    Dim xlApp As Excel.Application
    Dim wb As Excel.Workbook
    Dim ws As Excel.WorkSheet
    Dim strFilePath As String
    Dim tblComments As Excel.ListObject
    Dim intComment As Integer
    Dim intDetail As Integer

    'Initialise variables
    Set xlApp = CreateObject("Excel.Application")
    strFilePath = ActiveDocument.Path & "\" & Format(Str(Now), "yymmdd") & " Comments Register - " & Split(ActiveDocument.Name, ".")(0) & ".xlsx"
    
    With xlApp
        .Visible = False
        Set wb = .Workbooks.Add
        Set ws = wb.Worksheets(1)
        
        ' Set up header row
        ws.Cells(intHeaderRow, "A").Value = "Heading"
        ws.Cells(intHeaderRow, "B").Value = "Commented text"
        ws.Cells(intHeaderRow, "C").Value = "Comment"
        ws.Cells(intHeaderRow, "D").Value = "Resolved"
        ws.Cells(intHeaderRow, "E").Value = "Originator"
        
        ' Create table
        Set tblComments = ws.ListObjects.Add(xlSrcRange, ws.range(ws.Cells(intHeaderRow, "A"), ws.Cells(intHeaderRow + 1, "E")), , xlYes)
        tblComments.Name = "tblComments"
        
        For intComment = 1 To UBound(arrComments, 1)
            For intDetail = 1 To UBound(arrComments, 2)
                ws.Cells(intComment + intHeaderRow, intDetail).Value = arrComments(intComment, intDetail)
            Next intDetail
        Next intComment
        
        ' Format table
        ws.Columns("A:D").EntireColumn.AutoFit
        
        ' Set Excel application to be visible
        .Visible = True
        
        wb.SaveAs strFilePath
    
        Set ws = Nothing
        Set wb = Nothing
        Set xlApp = Nothing
    End With
End Sub