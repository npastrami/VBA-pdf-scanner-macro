Attribute VB_Name = "Module1"
Option Explicit

Dim pdfApp As Acrobat.AcroApp
Dim pdfDoc As Acrobat.AcroPDDoc
Dim avDoc As Acrobat.AcroAVDoc

Sub Get_PDFGetter()
    Dim wbActive As Workbook
    Dim wsHome As Worksheet, wsData As Worksheet, wsResults As Worksheet, wsList As Worksheet
    Dim pdfBookmarks As Object
    Dim bookmarkCount As Integer
    Dim maxBookmarkCount As Integer
    Dim column_Files As Range
    Dim DirLookup As Object
    Dim lastRow As Integer

    Dim i As Long
    Dim xRg As Range
    Dim xStr As String
    Dim xFd As FileDialog
    Dim xFdItem As Variant
    Dim xFileName As String
    Dim xFileNum As Long
    Dim RegExp As Object

    Dim iCounter As Integer, iCounter_List As Integer
    Dim iFileName As String, iFolderName As String
    Dim isResults As Boolean
    Dim iPgCount As Long

    Set wbActive = ActiveWorkbook
    Set wsHome = wbActive.ActiveSheet
    Set pdfApp = CreateObject("AcroExch.App")
    Set pdfDoc = CreateObject("AcroExch.PDDoc")
    Set DirLookup = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    Set wsData = wbActive.Worksheets("data")
    On Error GoTo 0

    If wsData Is Nothing Then Exit Sub

    With wsData
        Set column_Files = .Columns(3)
        lastRow = .Cells(.Rows.Count, column_Files.Column).End(xlUp).Row
    End With

    Application.DisplayAlerts = False

    If Evaluate("=IsRef('Folder Results'!A1)") Then
        Worksheets("Folder Results").Delete
    End If

    If Evaluate("=IsRef('Results'!A1)") Then
        Worksheets("Results").Delete
    End If

    Application.DisplayAlerts = True

    wbActive.Sheets.Add
    Set wsResults = ActiveSheet
    With wsResults
        .Name = "Folder Results"
        .Cells(1, 1).Value = "Folder List"
        .Cells(1, 2).Value = "PDF Page Count"
        .Cells(1, 3).Value = "Bookmark Count"
    End With

    wbActive.Sheets.Add
    Set wsList = ActiveSheet

    With wsList
        .Name = "Results"
        .Cells(1, 1).Value = "Filepath"
        .Cells(1, 2).Value = "PDF Page Count"
        .Cells(1, 3).Value = "Bookmark Count"
    End With

    Range(wsData.Cells(2, column_Files.Column), wsData.Cells(lastRow, column_Files.Column)).Copy
    wsResults.Cells(2, 1).PasteSpecial xlPasteValues

    iCounter_List = 2
    For iCounter = 2 To lastRow
        iFolderName = wsData.Cells(iCounter, column_Files.Column) & "\"

        If DirLookup.FolderExists(iFolderName) = False Or Len(iFolderName) < 2 Then
            wsResults.Cells(iCounter, 2).Value = "N/A"
            wsResults.Cells(iCounter, 3).Value = 0
            GoTo NextFolder
        End If

        iFileName = Dir(iFolderName)

        Set pdfApp = CreateObject("AcroExch.App")
        pdfApp.Show ' Make the application visible

        ' Create a new AVDoc object
        Set avDoc = CreateObject("AcroExch.AVDoc")

        ' Open the PDF document with error handling
        If avDoc.Open(iFolderName & iFileName, "") Then
            ' Get the PDDoc object from the AVDoc
            Set pdfDoc = avDoc.GetPDDoc()

            ' Check if the PDDoc object was retrieved successfully
            If Not pdfDoc Is Nothing Then
                Dim jsObject As Object
                Set jsObject = pdfDoc.GetJSObject()

                ' Extract bookmark count
                bookmarkCount = GetBookmarkCountFromJSObject(jsObject)

                ' Update the max bookmark count if necessary
                If bookmarkCount > maxBookmarkCount Then
                    maxBookmarkCount = bookmarkCount
                End If

                ' Close the PDF document
                pdfDoc.Close
            End If

            ' Close the AVDoc
            avDoc.Close (0)
        End If

        ' Release the objects
        Set pdfDoc = Nothing
        Set avDoc = Nothing
        pdfApp.Exit ' Exit the Acrobat application

        ' Print the file details and bookmark count in the "Folder Results" tab
        wsResults.Cells(iCounter, 2).Value = maxBookmarkCount

        ' Reset the max bookmark count
        maxBookmarkCount = 0

        iFileName = Dir()
        Do Until iFileName = vbNullString
            If InStr(1, UCase(iFileName), "1040") > 0 And _
               InStr(1, UCase(iFileName), "EXTENSION") = 0 And _
               InStr(1, UCase(iFileName), ".ZIP") = 0 And _
               InStr(1, UCase(iFileName), "SIGNED") = 0 And _
               InStr(1, UCase(iFileName), ".PDF") > 0 Then

                wsList.Cells(iCounter_List, 1).Value = iFolderName & iFileName

                Set RegExp = CreateObject("VBScript.RegExp")
                RegExp.Global = True
                RegExp.Pattern = "/Type\s*/Page[^s]"
                xFileNum = FreeFile
                Open (iFolderName & iFileName) For Binary As #xFileNum
                xStr = Space(LOF(xFileNum))
                Get #xFileNum, , xStr
                Close #xFileNum
                wsList.Cells(iCounter_List, 2).Value = RegExp.Execute(xStr).Count
                iCounter_List = iCounter_List + 1

                iPgCount = RegExp.Execute(xStr).Count
                With wsResults.Cells(iCounter, 2)
                    If .Value = vbNullString Then
                        .Value = iPgCount
                    Else
                        If CInt(.Value) < iPgCount Then .Value = iPgCount
                    End If
                End With

                ' Extract bookmark count based on criteria
                bookmarkCount = GetBookmarkCountFromCriteria(iFolderName & iFileName)

                ' Update the max bookmark count if necessary
                If bookmarkCount > maxBookmarkCount Then
                    maxBookmarkCount = bookmarkCount
                End If
            End If
            iFileName = Dir()
        Loop

        ' Print the highest bookmark count in the "Folder Results" tab
        wsResults.Cells(iCounter, 3).Value = maxBookmarkCount

NextFolder:
        isResults = False
        iFolderName = vbNullString
    Next iCounter

    ' Auto-fit columns in both sheets
    wsList.Columns.AutoFit
    wsResults.Columns.AutoFit
End Sub

Function GetBookmarkCountFromJSObject(jsObject As Object) As Integer
    Dim bookmarkCount As Integer
    Dim bookmarkRegExp As Object
    Dim matches As Object
    Dim match As Object
    
    ' Create the RegExp object to match bookmark names
    Set bookmarkRegExp = CreateObject("VBScript.RegExp")
    bookmarkRegExp.Pattern = "name:\s*""([^""]*)"""
    bookmarkRegExp.Global = True
    
    ' Extract the bookmark names using RegExp
    Set matches = bookmarkRegExp.Execute(jsObject.toSource())
    
    ' Check if there are any matches
    If matches.Count > 0 Then
        Dim bookmarkNames As Object
        Set bookmarkNames = CreateObject("Scripting.Dictionary")
        
        For Each match In matches
            ' Check if the bookmark name matches the criteria
            If InStr(1, match.SubMatches(0), "Alabama", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Alaska", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Arizona", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Arkansas", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "California", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Colorado", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Connecticut", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Delaware", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Florida", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Georgia", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Hawaii", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Idaho", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Illinois", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Indiana", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Iowa", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Kansas", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Kentucky", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Louisiana", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Maine", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Maryland", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Massachusetts", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Michigan", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Minnesota", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Mississippi", vbTextCompare) > 0 Or _
            InStr(1, match.SubMatches(0), "Missouri", vbTextCompare) > 0 Then
                ' Add the bookmark name to the dictionary to avoid duplicates
                bookmarkNames(match.SubMatches(0)) = True
            End If
        Next match
        
        ' Count the unique bookmark names
        bookmarkCount = bookmarkNames.Count
    End If
    
    ' Return the bookmark count
    GetBookmarkCountFromJSObject = bookmarkCount
End Function

Function GetBookmarkCountFromCriteria(filePath As String) As Integer
    Dim bookmarkCount As Integer
    
    ' Open the PDF file with Acrobat
    Set pdfApp = CreateObject("AcroExch.App")
    pdfApp.Show ' Make the application visible

    ' Create a new AVDoc object
    Set avDoc = CreateObject("AcroExch.AVDoc")
    
    ' Open the PDF document with error handling
    If avDoc.Open(filePath, "") Then
        ' Get the PDDoc object from the AVDoc
        Set pdfDoc = avDoc.GetPDDoc()
        
        ' Check if the PDDoc object was retrieved successfully
        If Not pdfDoc Is Nothing Then
            Dim jsObject As Object
            Set jsObject = pdfDoc.GetJSObject()
            
            ' Extract bookmark count based on criteria
            bookmarkCount = GetBookmarkCountFromJSObject(jsObject)
            
            ' Close the PDF document
            pdfDoc.Close
        End If
        
        ' Close the AVDoc
        avDoc.Close (0)
    End If
    
    ' Release the objects
    Set pdfDoc = Nothing
    Set avDoc = Nothing
    pdfApp.Exit ' Exit the Acrobat application
    
    ' Return the bookmark count
    GetBookmarkCountFromCriteria = bookmarkCount
End Function

