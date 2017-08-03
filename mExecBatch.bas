Attribute VB_Name = "mExecBatch"
' @(h) mExecBatch.BAS ver 1.0 ( 2017.08.01 Numtech)

' @(s)
'
' This module calls NIC batch related class modules.
' It corresponds to the "Execute Batch" button in worksheet "Settings".
'
'----------------------------------------------------------------------------------------
' The followings describe the module.
'
' Workflow of Main_ExecBatch()
'   - Start logging the process
'   - Check user input
'   - Get base data from calculation server
'   - Modify calculation setting file according to worksheet "Settings"
'   - Read replaced files directory
'   - Generate workspace input data by combining base data and replaced files (e.g., Tran-02.txt) for different scenarios
'   - Upload files, check data / run simulation, and download output using nicmd. nicmd is a command line tool that controls NtInsight products through a command prompt
'   - Open the files, retrieve necessary information (e.g., Marginal VaR) to Excel sheet, and consolidate the result
'   - Export consolidated results to txt file
'
'
' Below is a brief description of the roles of the modules and class modules used by this main function
'       mExecBatch - main(), execute the full batch process, corresponding to button "Execute Batch"
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
' Function    : Main() of button "Execute Batch" in worksheet "Settings"
'
' Return        :
'
' Argument :
'
' Feature     :
'
' Remarks   : It will execute the batch process completely

Public Sub Main_ExecBatch()

    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False

    Call logger.OutputInfo("Main_ExecBatch:Start")
    
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
    If Not nicBatch.NicmdHelper(False) Then
       GoTo ErrorHandler
    End If
    
    ' Parse the result to the result sheet, named as [FileCategory]-[AttributeName] (e.g., StatCFV-MVAR1)
    Dim nicExport As New clsExport
    If Not nicExport.ExportHelper() Then
        GoTo ErrorHandler
    End If
    
    ' Pop up a success window before exist
    MsgBox "Batch calculation finished. Please check results.", vbOKOnly + vbInformation
    
    GoTo Exit_main

ErrorHandler:
    Call logger.OutputError("Main_ExecBatch", Err)
    MsgBox "Error occured. Please check the log information", vbOKOnly + vbCritical
    GoTo Exit_main

Exit_main:
    ' Delete the working folders
    ' Result folder (i.e., raw output) will be maintained
    Call DeleteDir(g_baseWsDataDirPath)
    Call DeleteDir(g_tempDirPath)
    Call logger.OutputInfo("Main_ExecBatch:End")
    Application.ScreenUpdating = True
    End     ' End process, clear data in memory
End Sub
