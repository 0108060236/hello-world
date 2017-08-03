Attribute VB_Name = "mClear"
' @(h) mClear.BAS ver 1.0 ( 2017.08.01 Numtech)

' @(s)
'
' This module clears existing results from Excel sheets.
' It corresponds to the "Clear Result" button in worksheet "Settings".
'
'----------------------------------------------------------------------------------------
' The followings describe the module.
'
' Workflow of Main_Clear()
'   - Start logging the process
'   - Delete all the existing consolidated results
'
'
' Below is a brief description of the roles of the modules and class modules used by this main function
'       mGlobals - module that declares all the gobal constants, variables, and objects
'       mUtility - provide utility functions for other modules and class to use
'       clsLog - track the macro and identify errors
'
'----------------------------------------------------------------------------------------
Option Explicit

'
' Function    : Main() of button "Clear Result" in worksheet "Settings"
'
' Return       :
'
' Argument :
'
' Feature     :
'
' Remarks   :

Public Sub Main_Clear()

    Application.ScreenUpdating = False

    Call logger.OutputInfo("Main_ClearResult:Start")

    ' Delete all the result sheets before start
    If Not DeleteResultSheets Then
        GoTo ErrorHandler
    End If

    GoTo Exit_main
    
ErrorHandler:
    Call logger.OutputError("Main_ClearResult", Err)
    MsgBox "Error occured. Please check the log information", vbOKOnly + vbCritical
    GoTo Exit_main
    
Exit_main:
    Call logger.OutputInfo("Main_ClearResult:End")
    Application.ScreenUpdating = True
    End     ' End process, clear data in memory
End Sub
