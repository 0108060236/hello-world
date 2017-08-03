VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @(h) clsFileOutput.cls ver 1.0 ( 2017.08.01 Numtech)

' @(s)
'
' This class stores information and methods for exporting and consolidating txt output files

Option Explicit

Public fileCategory As String
Public fileHeader As String
Public attributeName As String
Public fileNames As Collection

Private Enum colType
    e_PkAndAttribute = 1
    e_AttributeOnly = 2
End Enum

'
' Function   : Function to perform all the tasks needed to be done for exporting txt files of a fileCategory (e.g., StatCFV of batchId = 1, 2, 3, etc.)
'
' Return      : A boolean that indicates whether the entire process is successful
'
' Argument :
'
' Feature     :
'
' Remarks    :

Public Function GenerateOutput() As Boolean

    GenerateOutput = False
    
    InitializeWorksheet  ' Initialize and create named worksheet
    
    Dim fileNames As String
    Dim isFirstFile As Boolean
    Dim vntRowData As Variant
    Dim arrData() As Variant
    Dim startColIndex As Integer
    Dim batchIdCounter As Integer
    fileNames = Dir(g_resultDirPath & fileCategory & "*.txt")
    isFirstFile = True ' A boolean that indicates whether it is the first file accessed; if so, primary key columns will be retrieved along with the target column attribute
    startColIndex = 1
    batchIdCounter = 1
    
    Do While Len(fileNames) > 0
        ' Load all the data from txt file
        vntRowData = LoadData(g_resultDirPath & fileNames)
        If isFirstFile = True Then
            ' When accessing a file for the first time
            arrData = GetColumns(vntRowData, e_PkAndAttribute, batchIdCounter)
            CopyToRng arrData, startColIndex
            startColIndex = startColIndex + UBound(arrData, 2) + 1
            isFirstFile = False
            batchIdCounter = batchIdCounter + 1
        Else
            ' When accessing the second time and onwards
            arrData = GetColumns(vntRowData, e_AttributeOnly, batchIdCounter)
            CopyToRng arrData, startColIndex
            startColIndex = startColIndex + UBound(arrData, 2) + 1
            batchIdCounter = batchIdCounter + 1
        End If
        fileNames = Dir     ' Loop to next file
    Loop
    
    FormatSheet     ' Format result sheet
    
    GenerateOutput = True
    
End Function

'
' Function   : Function to check if all the specified output attributes exist in the corresponding file.
'
' Return      : A boolean that indicates whether the entire process is successful
'
' Argument :
'
' Feature     :
'
' Remarks    : This function is used only when performing data check procedure

Public Function CheckIfAttributeExists() As Boolean
    
    CheckIfAttributeExists = False
    
    Dim fileNames As String
    Dim aFileHeader() As String
    Dim item As Variant
    fileNames = Dir(g_resultDirPath & fileCategory & "*.txt")
    aFileHeader = Split(fileHeader, vbTab)
    
    For Each item In aFileHeader
        If item = attributeName Then
            CheckIfAttributeExists = True
            Exit Function
        End If
    Next item
    
End Function

'
' Function   : Create new worksheet for consolidating results
'
' Return      : A boolean that indicates whether the creation process is successful
'
' Argument :
'
' Feature     :
'
' Remarks    : Each result worksheet contains batch results of a single column
'                         The worksheet is named as fileCategory - attributeName (e.g., StatCFV-MVar1)

Private Function InitializeWorksheet() As Boolean

    ' Create new worksheet for the output
    Dim ws As Worksheet
    With ThisWorkbook
        Set ws = .Worksheets.Add(after:=.Sheets(.Sheets.Count))
        ws.Name = fileCategory & "-" & attributeName
    End With

End Function

'
' Function   : Read txt file and return the raw data as row of arrays
'
' Return      : Arrays of row data
'
' Argument : statFilePath - the path of the output txt file
'
' Feature    :
'
' Remarks   :

Private Function LoadData(ByVal statFilePath As String) As Variant
    
    Dim fileNumber As Integer
    Dim dataString As String
    fileNumber = freeFile   ' Obtain next available file number
    
    Open statFilePath For Binary As #fileNumber
    dataString = Space$(LOF(fileNumber))    ' Sets buffer to Length Of File
    Get #fileNumber, , dataString   ' Fits exactly
    Close #fileNumber   ' Close file
    
    LoadData = Split(dataString, vbNewLine)
    
End Function

'
' Function    : Get the data of the specified columns
'
' Return       : A 2 dimensional array that stores the data of a specified column
'
' Argument  : vntRowData - data loaded from a txt output file
'                         colType - type of column data
'                         batchIdCounter - a counter to add batchId to attribute
'
' Feature     :
'
' Remarks    :

Private Function GetColumns(ByRef vntRowData As Variant, ByVal colType As colType, ByVal batchIdCounter As Integer) As Variant()
    
    Dim colIndexes() As Integer     ' Indexes of the specified column
    Dim arr() As Variant    ' 2-dimensional array to save data
    Dim i, j As Integer
    
    ' Get column indexes in an array according to the column type
    Select Case colType
    
        ' When accessing an output file the first time, both the primary keys and specified attribute should be added to the indexes
        Case e_PkAndAttribute
            colIndexes = GetPkColIndexes    ' Indexes of primary keys
            ReDim Preserve colIndexes(UBound(colIndexes) + 1) As Integer
            colIndexes(UBound(colIndexes)) = GetColIndex    ' Add specified attribute
            
        ' When accessing an output file the second time and onwards, only the specified attribute needs to be added to the index
        ' To make codes re-usable, a single attribute index is also saved in an array
        Case e_AttributeOnly
            ReDim colIndexes(0) As Integer
            colIndexes(0) = GetColIndex
            
    End Select
    
    ReDim arr(0 To UBound(vntRowData) - 1, 0 To UBound(colIndexes)) As Variant
    
    ' NtInsight produces the last blank row
    For i = LBound(vntRowData) To UBound(vntRowData) - 1
        For j = LBound(colIndexes) To UBound(colIndexes)
            arr(i, j) = Split(vntRowData(i), vbTab)(colIndexes(j))
        Next j
    Next i
    
    ' Add batchId to distinguish results from different batches
    arr(0, UBound(arr, 2)) = arr(0, UBound(arr, 2)) & "(Batch " & batchIdCounter & ")"
    
    GetColumns = arr

End Function

'
' Function    : Aquire the index / position of a specific attributeName (i.e., which column the attribute is at)
'
' Return        : A integer that indicates the column index of the attribute within a particular output txt file
'
' Argument  :
'
' Feature      :
'
' Remarks    :

Private Function GetColIndex() As Integer

    Dim i As Integer
    Dim aFileHeader() As String
    ' Split fileHeader string as array
    aFileHeader = Split(fileHeader, vbTab)

    For i = LBound(aFileHeader) To UBound(aFileHeader)
        If aFileHeader(i) = attributeName Then
            GetColIndex = i
        End If
    Next i
    
End Function

'
' Function   : Get the number of primary keys in the output file
'
' Return       : An array of column indexes for primary keys (including the name attributes)
'
' Argument  :
'
' Feature     : The column indexes include both primary key and its corresponding names (e.g., both TLayerX and TLayerName are included)
'
' Remarks   :

Private Function GetPkColIndexes() As Integer()
    
    Dim matches As String
    Dim noOfCols As Integer
    Dim i As Integer
    Dim pkColIndexes() As Integer
    
    ' Find all the matches for primary key columns
    matches = RegularExpression(fileHeader, ".Layer.*?\t")
    noOfCols = UBound(Split(matches, vbTab & " "))     ' largest subscript, function RegularExpression split each element by a space

    ReDim pkColIndexes(noOfCols) As Integer
    
    For i = LBound(pkColIndexes) To UBound(pkColIndexes)
            pkColIndexes(i) = i
    Next i
    
    GetPkColIndexes = pkColIndexes
    
End Function

'
' Function    : Copy a 2 dimensional array to a range
'
' Return       :
'
' Argument  : arrData - 2 dimensional consolidated data
'                         startColIndex - indicates where / which colukmn the data should be appended
'
' Feature     : Worksheet is chosen by standard worksheet name, fileCategory-attributeName
'
' Remarks    :

Private Function CopyToRng(ByRef arrData() As Variant, ByVal startColIndex As Integer) As Boolean
    
    ' These parameter will be used to define the range to copy
    Dim noOfRecords As Integer
    Dim noOfCols As Integer
    noOfRecords = UBound(arrData, 1) - LBound(arrData, 1) + 1
    noOfCols = UBound(arrData, 2) - LBound(arrData, 1) + 1

    With Sheets(fileCategory & "-" & attributeName)
        .Range(.Cells(1, startColIndex), .Cells(noOfRecords, startColIndex + noOfCols - 1)).Value = arrData
    End With

End Function

'
' Function   : Format the result sheet
'
' Return       : A boolean value that indicates whether the operation is successful
'
' Argument  :
'
' Feature     :
'
' Remarks   :
Private Function FormatSheet() As Boolean
    
    FormatSheet = False
    
    With Sheets(fileCategory & "-" & attributeName)
        .Rows(1).Font.Bold = True
        .Range("A1").CurrentRegion.Rows(1).Interior.Color = RGB(214, 220, 220)
        .Range("A1").CurrentRegion.ColumnWidth = 25
    End With

    FormatSheet = False
    
End Function

