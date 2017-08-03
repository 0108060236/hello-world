VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @(h) clsValidator.cls ver 1.0 ( 2017.08.01 Numtech)

' @(s)
'
' This class performs all the user input checking (e.g., worksheet "BatchDetail")
' Each checking function for a column in a worksheet is named as [ParentWorksheet]_[ChildWorksheet(if necessary)]_[AttributeName]

Option Explicit

'
' Function   : Wrapper function to perform all the data checkings
'
' Return      : A boolean value that indicates whether the checking process is executed succesfully
'
' Argument : None
'
' Feature      :
'
' Remarks    :

Public Function Validate() As Boolean
    
    ' Write to log about the progression
    Call logger.OutputInfo("Validation: checking user inputs...")
    
    Validate = Settings_PATH_NICMD And Settings_DIR_OUTPUT And BatchDetail_BatchId And BatchMaster_BatchId And BatchDetail_FileName And BatchMaster_BatchDetail_BatchId
    
    Call logger.OutputInfo("Validation: success.")
    
End Function

'
' Function   : Check if nicmd.exe is in the specified directory
'
' Return      : Boolean value to indicate whether nicmd.exe is in the specified directory
'
' Argument : None
'
' Feature     :
'
' Remarks   :

Private Function Settings_PATH_NICMD()
    
    Settings_PATH_NICMD = False
    
    ' Check if nicmd exists
    If Dir(g_pathNicmd, vbDirectory) = "" Then
        Call logger.OutputError("nicmd.exe does not exist in the specified directory. Please check settings (PATH_NICMD).", Err)
        Exit Function
    End If
    
    Settings_PATH_NICMD = True
    
End Function

'
' Function   : Check if output file exists
'
' Return      : Boolean value to indicate whether the folder exists
'
' Argument : None
'
' Feature     :
'
' Remarks   :

Private Function Settings_DIR_OUTPUT()
    
    Settings_DIR_OUTPUT = False
    
    ' Check if nicmd exists
    If Dir(g_txtOutputDirPath, vbDirectory) = "" Then
        Call logger.OutputError("Output folder does not exist in the specified directory. Please check settings (DIR_OUTPUT).", Err)
        Exit Function
    End If
    
    Settings_DIR_OUTPUT = True
    
End Function

'
' Function   : Check if batch number is entered correctly
'
' Return      : Boolean value to indicate whether the BatchIds are entered correctly
'
' Argument : None
'
' Feature     :
'
' Remarks   : The batchIds should be enterred as running number sequence starting from 1

Private Function BatchMaster_BatchId() As Boolean

    Dim isValid As Boolean
    Dim i As Integer
    Dim batchIdCell As Range
    Dim batchIdRng As Range

    isValid = True
    i = 1   ' A counter that helps to check running number constraint
    
    With wsBatchMaster
        ' Cells(2, 1) : Start of "BatchId" column
        Set batchIdRng = .Range(.Cells(2, 1), .Cells(2, 1).End(xlDown))
        For Each batchIdCell In batchIdRng
            If batchIdCell.Value <> i Then
                isValid = False
                Exit For
            End If
            i = i + 1
        Next batchIdCell
    End With
    
    If isValid = False Then
        Call logger.OutputError("The batch Id is not entered correctly. Please check worksheet ""BatchMaster"".", Err)
            Exit Function
    End If
    
    BatchMaster_BatchId = isValid
    
End Function

'
' Function   : Check if batch number is entered correctly
'
' Return      : Boolean value to indicate whether the BatchIds are entered correctly
'
' Argument : None
'
' Feature     :
'
' Remarks   : The batchIds should be enterred as a non-decreasing integer sequence starting from 1
'                        Column E in worksheet "BatchDetail" is a hidden column that helps validate the validity of each batch Id
'                        The checking column E checks only till row 200

Private Function BatchDetail_BatchId() As Boolean
    
    Dim isValid As Boolean
    Dim hasErrorCell As Range
    Dim hasErrorRng As Range
    
    isValid = True
    
    With wsBatchDetail
        ' Cells(2, 5) : Start of "HASERROR" column
        Set hasErrorRng = .Range(.Cells(2, 5), .Cells(2, 5).End(xlDown))
    End With
    
    For Each hasErrorCell In hasErrorRng
        If hasErrorCell.Value = True Then
            isValid = False
            Exit For
        End If
    Next hasErrorCell
    
    ' Check if the batchId is entered as as a non-decreasing integer sequence starting from 1
    If isValid = False Then
        Call logger.OutputError("The batch Id is not entered correctly. Please check worksheet ""BatchDetail"".", Err)
            Exit Function
    End If
    
    BatchDetail_BatchId = isValid

End Function

'
' Function   : Check if all the replace files specified are in the specified directory
'
' Return      : A boolean value that indicates whether the checking process is successful
'
' Argument : None
'
' Feature      :
'
' Remarks    : One side checking only (ensure the files in the input are in the folder directory)

Private Function BatchDetail_FileName() As Boolean

    BatchDetail_FileName = False
    
    Dim fileNameCell As Range
    Dim fileNameRng As Range
    Dim fullFilePath As String
    
    With wsBatchDetail
        ' Get the full input range of file names
        ' Cells(2, 3) : Start of "FileName" column
        Set fileNameRng = .Range(.Cells(2, 3), .Cells(2, 3).End(xlDown))
    End With
    
    For Each fileNameCell In fileNameRng
        fullFilePath = g_replaceDirPath & "\" & fileNameCell.Text & ".txt"
        If Dir(fullFilePath, vbDirectory) = "" Then
            Call logger.OutputError("The specified overwritten files do not exist in the local folder. Please check settings overwritten files.", Err)
            Exit Function
        End If
    Next fileNameCell

    BatchDetail_FileName = True
    
End Function

'
' Function   : This function will check if all the batchIds specified in "BatchMaster" are also specified in "BatchDetail"
'
' Return      : Boolean value to indicate whether the BatchIds are entered correctly
'
' Argument : None
'
' Feature     :
'
' Remarks   :

Private Function BatchMaster_BatchDetail_BatchId() As Boolean

    BatchMaster_BatchDetail_BatchId = False
    
    Dim BatchMaster_NoOfBatch As Integer
    Dim BatchDetail_NoOfBatch As Integer
    
    With wsBatchMaster
        ' Cells(2, 1) : Start of "BatchId" column
        BatchMaster_NoOfBatch = .Cells(2, 1).End(xlDown).row - 1
    End With
    
    With wsBatchDetail
        ' Cells(2, 1) : Start of "BatchId" column
        BatchDetail_NoOfBatch = CountUnique(.Range(.Cells(2, 1), .Cells(2, 1).End(xlDown)))
    End With
    
    If BatchMaster_NoOfBatch <> BatchDetail_NoOfBatch Then
            Call logger.OutputError("The number of batches in worksheet ""BatchMaster"" and ""BatchDetail"" do not match. Please check the two sheets.", Err)
            Exit Function
    End If
    
    BatchMaster_BatchDetail_BatchId = True
        
End Function

