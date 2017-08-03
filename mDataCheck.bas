Attribute VB_Name = "mDataCheck"
' @(h) mDataCheck.BAS ver 1.0 ( 2017.08.01 Numtech)

' @(s)
'
' This module calls data check only related modules.
' It corresponds to the "Data Check" button in worksheet "Settings".
'
'----------------------------------------------------------------------------------------
' The followings describe the module.
'
' Workflow of Main_DataCheck()
'   - Start logging the process
'   - Check user input
'   - Get base data from calculation server
'   - Modify calculation setting file and set the number of iterations as minimum 10 for all batches
'   - Generate workspace input data by combining base data and replaced files (e.g., Tran-02.txt) for different scenarios
'   - Upload files, check data / run simulation, and download output using nicmd. nicmd is a command line tool that controls NtInsight products through a command prompt
'   - Open the files, check if all the specified attributes and output files are retrievable
'
'
' Below is a brief description of the roles of the modules and class modules used by this main function
'       mDataCheck - main(), perform data check process, corresponding to button "Data Check"
'       mGlobals - module that declares all the gobal constants, variables, and objects
'       mUtility - provide utility functions for other modules and class to use
'       clsLog - track the macro and identify errors
'       clsValidator - validate the user inputs before calculation starts
'       clsNicmd - nicmd basic commands and orders
'       clsCalcConfig - read, modify and save calculation setting file (e.g., Cb.xml)
'                                  - the data is read from worksheet "BatchMaster"
'       clsCalcConfigs - typed collection of clsCalcConfig object, including methods assciated with collection of clsCalcConfig objects
'       clsBatchDetail - store detail information for each batch
'                                    - the data is read from worksheet "BatchDetail"
'       clsBatchDetails - typed collection of clsBatchDetail object, including methods assciated with collection of clsBatchDetail objects
'       clsBatch  - perform nicmd commands in a batch sequence
'       clsFileOutput - generate consolidated output in Excel worksheet
'       clsFileOutputs - typed collection of clsFileOutput object, including methods assciated with collection of clsFileOutput objects
'       clsExport  - wrap tasks in clsFileOutput, and export the consolidated result into txt files
'
'----------------------------------------------------------------------------------------
Option Explicit

'
' Function   : Main() of button "Data Check" in worksheet "Settings"
'
' Return       :
'
' Argument :
'
' Feature     :
'
' Remarks   :

Public Sub Main_DataCheck()

    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False

    Call logger.OutputInfo("Main_DataCheck:Start")
    
    ' Delete all the result sheets before start
    If Not DeleteResultSheets Then
        GoTo ErrorHandler
    End If
    
    ' Check validity of user input data
    Dim validator As New clsValidator
    If Not validator.Validate() Then
        GoTo ErrorHandler
    End If
    
    ' Temporary folder to store base workspace data
    Call CreateDir(g_baseWsDataDirPath, "")
    Call CreateDir(g_tempDirPath, "")
    Call CreateDir(g_resultDirPath, "")

    ' Batch process is performed here: download base data, modify base data, replace files, run simulation, download result
    Dim nicBatch As New clsBatch
    ' NicmdHelper takes one boolean argument: True -> Data check process; False -> Batch process
    If Not nicBatch.NicmdHelper(True) Then
        GoTo ErrorHandler
    End If
    
    Dim nicExport As New clsExport
    ' Data check of exporting files
    If Not nicExport.ExportHelper_DataCheck() Then
        GoTo ErrorHandler
    End If
    
    ' Pop up a success window before exist
    MsgBox "Data check succeed. Do not forget to modify setting before proceeding to the batch calculation.", vbOKOnly + vbInformation
    
    GoTo Exit_main

ErrorHandler:
    Call logger.OutputError("Main_DataCheck", Err)
    MsgBox "Data check failed. Please check the log information", vbOKOnly + vbCritical
    GoTo Exit_main

Exit_main:
    ' Delete the working folders
    ' Result folder (i.e., raw output) will not be maintained
    Call DeleteDir(g_baseWsDataDirPath)
    Call DeleteDir(g_tempDirPath)
    Call DeleteDir(g_resultDirPath)
    Call logger.OutputInfo("Main_DataCheck:End")
    Application.ScreenUpdating = True
    End     ' End process, clear data in memory
End Sub

