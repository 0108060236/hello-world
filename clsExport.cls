VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @(h) clsExport.cls ver 1.0 ( 2017.08.01 Numtech)

' @(s)
'
' This class processes the NIC result file (e.g., StatCFV)

Option Explicit

Private fileOutputs As clsFileOutputs

'
' Function   : Initialize object
'
' Return      : None
'
' Argument : None
'
' Feature     :
'
' Remarks   :
Private Sub Class_Initialize()
    
End Sub

'
' Function   : Helper function to execute all the tasks for wrangling and exporting results
'
' Return      : A boolean value that indicates whether the output process is successful
'
' Argument : None
'
' Feature     :
'
' Remarks   :

Public Function ExportHelper() As Boolean

    ExportHelper = False
    
    Dim i As Integer
    
    Init_fileOutputs    ' Initialize and create objects of clsFileOutputs

    For i = 1 To fileOutputs.Count
        Call logger.OutputInfo("Generate output for " & fileOutputs.item(i).fileCategory & "-" & fileOutputs.item(i).attributeName)
        fileOutputs.item(i).GenerateOutput     ' Output consolidated results to Excel worksheet
        Call logger.OutputInfo("Output generated for " & fileOutputs.item(i).fileCategory & "-" & fileOutputs.item(i).attributeName)
    Next i
    
    Call logger.OutputInfo("Exporting results to txt files...")
    ExportTxtFiles  ' Export consolidated results to txt file
    Call logger.OutputInfo("Exporting successful.")
    
    ExportHelper = True
    
End Function

'
' Function   : Helper function that performs data checking for exporting process
'
' Return      : A boolean value that indicates whether the output process is successful
'
' Argument : None
'
' Feature     :
'
' Remarks   :

Public Function ExportHelper_DataCheck() As Boolean

    ExportHelper_DataCheck = False
    
    Dim i As Integer
    
    Init_fileOutputs    ' Initialize and create objects of clsFileOutputs

    For i = 1 To fileOutputs.Count
        Call logger.OutputInfo("Check if  " & fileOutputs.item(i).attributeName & " exists in " & fileOutputs.item(i).fileCategory & "...")
        ' Check if all the attributes exist in the specifiied output files
        If Not fileOutputs.item(i).CheckIfAttributeExists Then
            Call logger.OutputError("Specified attribute cannot be found in the corresponding fileCategory. Please check worksheet ""FileOutput"".", Err)
            Exit Function
        End If
        Call logger.OutputInfo("Checking succeed.")
    Next i
    
    ExportHelper_DataCheck = True
    
End Function

'
' Function   : Export the consolidated results stored in excel sheets to separate txt files
'
' Return      : An boolean that indicates whether the operation is sucessful
'
' Argument : None
'
' Feature     :
'
' Remarks   :

Private Function ExportTxtFiles() As Boolean

    ExportTxtFiles = False
    
    Dim i As Integer
    Dim filePath As String
    Dim ws As Worksheet
    Dim arr As Variant  ' Range variant to store data temporarily
    
    For Each ws In ActiveWorkbook.Sheets
        If InStr(g_staticSheetNames, ws.Name) = 0 Then
            With ws
                filePath = g_txtOutputDirPath & "\" & .Name & ".txt"
                If Dir(filePath) <> "" Then
                    Kill filePath
                End If
                Open g_txtOutputDirPath & "\" & .Name & ".txt" For Output As #1
                arr = .Range("A1").CurrentRegion
                For i = 1 To UBound(arr)
                    Print #1, Join(Application.Index(arr, i), vbTab)
                Next i
                Close #1
            End With
        End If
    Next ws
    
    ExportTxtFiles = True
    
End Function

'
' Function    : Initialize clsFileOutputs object by reading values from worksheet "wsFileOutput"
'
' Return        : A boolean that indicates whether the initialization is successful
'
' Argument : None
'
' Feature      :
'
' Remarks    :

Private Function Init_fileOutputs() As Boolean
    
    Init_fileOutputs = False
    
    Dim i As Integer
    Dim fileCategory As String
    Dim attributeName As String
    
    Set fileOutputs = New clsFileOutputs

    With wsFileOutput
        For i = 2 To .Range("A1").CurrentRegion.Rows.Count
            fileCategory = .Cells(i, 1).Value
            attributeName = .Cells(i, 2).Value
            fileOutputs.Add fileCategory, LoadFileHeader(fileCategory), attributeName, LoadFileNames(fileCategory)
        Next i
    End With
    
    Init_fileOutputs = True
    
End Function

'
' Function    : Load / Get fileHeader of a specified output fileCategory
'
' Return        : A string of fileHeader split by tab
'
' Argument : fileCategory - Output file category (standard name without ".txt")
'
' Feature      :
'
' Remarks    :

Private Function LoadFileHeader(ByVal fileCategory As String) As String

    Dim statFilePath1 As String
    
    ' The file path of the specified fileCategory (batchId = 1)
    statFilePath1 = g_resultDirPath & fileCategory & "-1.txt"
    
    With CreateObject("Scripting.FileSystemObject").OpenTextFile(statFilePath1)
        LoadFileHeader = .ReadLine
        .Close
    End With
    
End Function

'
' Function    : Load / Get fileNames of a specified output fileCategory
'
' Return        : fileNames stored in a collection; each fileName being an item in a collection
'
' Argument : fileCategory - Output file category (standard name without ".txt")
'
' Feature      :
'
' Remarks    :

Private Function LoadFileNames(ByVal fileCategory As String) As Collection
    
    Dim fileName As String
    Dim fileNames As Collection
    
    fileName = Dir(g_replaceDirPath & fileCategory & "*.txt")
    Set fileNames = New Collection
    
    Do While fileName <> ""
        fileNames.Add fileName
        fileName = Dir()
    Loop
    
    Set LoadFileNames = fileNames

End Function
