VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @(h) clsBatch.cls ver 1.0 ( 2017.08.01 Numtech)

' @(s)
'
' This class performs nicmd commands in a batch manner

Option Explicit

Private nicmd As clsNicmd
Private calcConfigs As clsCalcConfigs
Private batchDetails As clsBatchDetails

'
' Function     : This function uses nicmd to connect -> prepare for base data -> upload Tran -> execute simulation -> download result
'
' Return        : Boolean value that indicates whether the entire process is executed sucessfully
'
' Argument  : None
'
' Feature       : When upload, simulation, or dowloading process fails, the function continues from the next item.
'
' Remarks     :

Public Function NicmdHelper(ByVal isDataCheckOnly As Boolean) As Boolean

    NicmdHelper = False
    
    If nicmd Is Nothing Then
        Init_nicmd
    End If
    
    If calcConfigs Is Nothing Then
        Init_calcConfigs isDataCheckOnly
    End If
    
    If batchDetails Is Nothing Then
        Init_batchDetails
    End If
    
    ' Connect the server
    If Not nicmd.NicmdConfig() Then
        Exit Function
    End If
    
    Dim maxBatchNo  As Long
    Dim resultMsg  As String
    Dim batchId As Long
    maxBatchNo = calcConfigs.Count
    
    For batchId = 1 To maxBatchNo
        
        resultMsg = "Batch Id " & batchId & ": Batch processing started..."
        Call logger.OutputInfo(resultMsg)
        
        ' Update the Temp folder and prepare for calculation
        Call logger.OutputInfo("Updating working folder...")
        If Not UpdateWorkingFolder(batchId) Then
            Exit Function
        End If
        Call logger.OutputInfo("Working folder updated")
        
        If Not NicmdUpload Then
            resultMsg = "Batch Id " & batchId & ": Upload failed"
            GoTo Warning
        End If
        
        If Not NicmdResimulate Then
            resultMsg = "Batch Id " & batchId & ": Simulation failed"
            GoTo Warning
        End If
        
        If Not NicmdDownload(batchId) Then
            resultMsg = "Batch Id " & batchId & ": Download failed"
            GoTo Warning
        End If
        
        resultMsg = "Batch Id " & batchId & ": Batch processing succeed. (Number of iterations: " & calcConfigs.item(batchId).noOfSystemIterations & ")"
        Call logger.OutputInfo(resultMsg)
    
        GoTo NextRow
        
Warning:

    Call logger.OutputWarn(resultMsg)
    GoTo NextRow

NextRow:
    Next batchId
    
    NicmdHelper = True

End Function

'
' Function   : Update working folder, [Temp], with the base data file, then replace with replace files in [DIR_OVERWRITTENFILE]
'
' Return      : Boolean value that indicates whether updating replace files is successful
'
' Argument : None
'
' Feature      : All the replacable files will be updated with base data files for every set of simulation
'
' Remarks    : In this version, only Tran.txt and Obligor.txt can be replaced

Private Function UpdateWorkingFolder(ByVal batchId As Integer) As Boolean

    UpdateWorkingFolder = False
    
    Dim batchFileNames As Dictionary
    Set batchFileNames = New Dictionary
    Dim key As Variant
    Dim item As Variant
    Dim inputFileNames As String
    Set batchFileNames = batchDetails.GetBatchFileNames(batchId)
    
    If batchId = 1 Then
        inputFileNames = nicmd.GetFileNames(g_baseWorkspace).item("INPUT")
        ' Download xml and input txt files
        nicmd.Download g_baseWorkspace, g_baseWsDataDirPath, inputFileNames
    End If
    
    calcConfigs.item(batchId).ConfigCalcSetting
    
    ' Copy base files from BaseWorkspaceData folder
    For Each item In g_sptReplaceFileNames
        FileCopy g_baseWsDataDirPath & item, g_tempDirPath & item
    Next item
    FileCopy g_baseWsDataDirPath & "Cb.xml", g_tempDirPath & "Cb.xml"
    
    ' Update the files with the ones in [DIR_OVERWRITTENFILE]
    For Each key In batchFileNames.Keys
        FileCopy g_baseWsDataDirPath & key, g_tempDirPath & key
        FileCopy g_replaceDirPath & "\" & batchFileNames.item(key), g_tempDirPath & key
    Next key
    
    UpdateWorkingFolder = True
        
End Function

'
' Function           : Upload files through nicmd
'
' Return               : Boolean value that indicates whether the upload is successful.
'
' Argument         : None
'
' Feature              : Upload the files in the Temp folder
'
' Remarks            : For every set of simulation, the replace files are stored in Temp folder, and this function will upload to server, after which it will delete the file

Private Function NicmdUpload() As Boolean

    NicmdUpload = False
    
    Dim item As Variant
    
    With nicmd
        ' Upload the supported replace files and the updated calculation setting, Cb.xml
        If Not .Upload(g_targetWorkspace, g_tempDirPath, Join(g_sptReplaceFileNames, " ") & " Cb.xml") Then
            Exit Function
        End If
    End With
    
    For Each item In g_sptReplaceFileNames
        Kill g_tempDirPath & "\" & item
    Next item
    
    Kill g_tempDirPath & "\Cb.xml"
    
    NicmdUpload = True
    
End Function

'
' Function  : Run simulation through nicmd
'
' Return      : Boolean value that indicates whether the simulation is completed successfully
'
' Argument : None
'
' Feature     :
'
' Remarks    :

Private Function NicmdResimulate() As Boolean

    NicmdResimulate = False

    With nicmd
        If Not .Resimulate(g_targetWorkspace) Then
            Exit Function
        End If
        
        ' Error checking
        Dim logArray    As Variant
        Dim i   As Long
        Dim isCalcCompleted   As Boolean
        Dim errorCount  As Long
        logArray = .GetOutputStringAsArray
        errorCount = 0
        For i = 0 To UBound(logArray)
            If InStr(logArray(i), "ERROR:") <> 0 Then
                errorCount = errorCount + 1
            End If

            If InStr(logArray(i), "CALCULATION COMPLETED") <> 0 Then
                isCalcCompleted = True
            End If
        Next

        If Not (errorCount = 0 And isCalcCompleted) Then
            Exit Function
        End If
    End With
    
    NicmdResimulate = True
           
End Function

'
' Function   : Download files through nicmd
'
' Return      : Boolean value that indicates whether the download is sucessful
'
' Arugment : batchId - pattern number
'
' Feature     :
'
' Remarks   : Workspace files will be deleted in the end, and that is why files are renamed as Tran-0.txt

Private Function NicmdDownload(ByVal batchId As Long) As Boolean
   
    NicmdDownload = False
    
    Dim resultFileNames As Dictionary
    Dim resultFileName As Variant
    Dim resultFileNameString As String
    
    ' Get the file names to download stored in a dictionary
    With wsFileOutput
        Set resultFileNames = GetUniqueValues(.Range(.Cells(2, 1), .Cells(2, 1).End(xlDown)))
    End With

    ' Put all the items in one string split by space for executing through nicmd using one line (consistent with other commands, and also cleaner)
    For Each resultFileName In resultFileNames
        resultFileNameString = resultFileNameString & resultFileName & ".txt "
    Next resultFileName
    
    ' Download all the files using the string of names
    With nicmd
        If Not .Download(g_targetWorkspace, g_resultDirPath, resultFileNameString) Then
            Exit Function
        End If
    End With
    
    ' Give batch id to file name
    For Each resultFileName In resultFileNames
        Name g_resultDirPath & "\" & resultFileName & ".txt" As g_resultDirPath & "\" & resultFileName & "-" & batchId & ".txt"
    Next resultFileName
    
    NicmdDownload = True
    
End Function

'
' Function   : Initialize nicmd object
'
' Return       : Boolean value that indicates whether the initialization is successful
'
' Argument : None
'
' Feature      :
'
' Remarks    :

Private Function Init_nicmd() As Boolean
    
    Init_nicmd = False
    
    Set nicmd = New clsNicmd
    
    Init_nicmd = True
    
End Function

'
' Function   : Initialize calcConfigs object
'
' Return      : Boolean value that indicates whether the initialization is successful
'
' Argument : isDataCheckOnly - boolean that indicates whether the object is initialized for data check or calculation; if it is for data check only, minimum 10 iterations is used
'
' Feature      :
'
' Remarks    : It reads data  from worksheet "wsBatchMaster"

Private Function Init_calcConfigs(ByVal isDataCheckOnly As Boolean) As Boolean
    
    Init_calcConfigs = False
    
    Dim i As Integer
    Set calcConfigs = New clsCalcConfigs

    With wsBatchMaster
        For i = 2 To .Range("A1").CurrentRegion.Rows.Count
            Select Case isDataCheckOnly
                Case True
                    ' Minimum iteration of 10 is used for all batches for data check
                    calcConfigs.Add .Cells(i, 1).Value, .Cells(i, 2).Text, 10, CStr(.Cells(i, 1).Value)
                Case False
                    calcConfigs.Add .Cells(i, 1).Value, .Cells(i, 2).Text, .Cells(i, 3).Value, CStr(.Cells(i, 1).Value)
            End Select
        Next i
    End With

    Init_calcConfigs = True
    
End Function

'
' Function   : Initialize batchDetails object
'
' Return       : Boolean value that indicates whether the initialization is successful
'
' Argument :
'
' Feature      :
'
' Remarks    : It reads data  from worksheet "wsBatchDetail"

Private Function Init_batchDetails() As Boolean
    
    Init_batchDetails = False
    
    Dim i As Integer
    Set batchDetails = New clsBatchDetails

    With wsBatchDetail
        For i = 2 To .Range("A1").CurrentRegion.Rows.Count
            batchDetails.Add .Cells(i, 1).Value, .Cells(i, 2).Text, .Cells(i, 3).Text
        Next i
    End With

    Init_batchDetails = True
    
End Function
