Attribute VB_Name = "mGlobals"
' @(h) mGlobals.BAS ver 1.0 ( 2017.08.01 Numtech)

' @(s)
'
' This module includes all the global, "constant" variables

Option Explicit

' The object for logging information as txt file
Public logger As New clsLog

' [USER INPUT] File path of nicmd.exe
Public Property Get g_pathNicmd()
    g_pathNicmd = GetParamValue("PATH_NICMD")
End Property

' [USER INPUT] Replace files folder directory path, files in this folder will be used to replace base data files
Public Property Get g_replaceDirPath()
    g_replaceDirPath = GetParamValue("DIR_OVERWRITTENFILE")
End Property

' [USER INPUT] Output file folder directory to store consolidated txt files
Public Property Get g_txtOutputDirPath()
    g_txtOutputDirPath = GetParamValue("DIR_OUTPUT")
End Property

' [USER INPUT] Host name for nicmd connection
Public Property Get g_hostName()
    g_hostName = GetParamValue("HOST_NAME")
End Property

' [USER INPUT] Port number for nicmd connection
Public Property Get g_portNo()
    g_portNo = CLng(GetParamValue("PORT_NUMBER"))
End Property

' [USER INPUT] User name for nicmd connection
Public Property Get g_userName()
    g_userName = GetParamValue("USER_NAME")
End Property

' [USER INPUT] Base workspace, every replace files will be overwritten based on this set of data
Public Property Get g_baseWorkspace()
    g_baseWorkspace = GetParamValue("BASE_WORKSPACE")
End Property

' [USER INPUT] Target working workspace, every set of calculation will be performed in this workspace by overwriting the data
Public Property Get g_targetWorkspace()
    g_targetWorkspace = GetParamValue("TARGET_WORKSPACE")
End Property

' Temporary folder directory path for downloading base workspace data to local
Public Property Get g_baseWsDataDirPath()
    g_baseWsDataDirPath = ThisWorkbook.Path & "\BaseWorkspaceData\"
End Property

' Temporary folder directory, this will be used to store files needed to be uploaded for each batch, and all replaced files will be uploaded through this folder
Public Property Get g_tempDirPath()
    g_tempDirPath = ThisWorkbook.Path & "\Temp\"
End Property

' Temporary folder directory, this will be used to store raw output files from each set of calculation
Public Property Get g_resultDirPath()
    g_resultDirPath = ThisWorkbook.Path & "\Result\"
End Property

' All the worksheet names that should not be deleted or refreshed among batch sequences
Public Property Get g_staticSheetNames()
    g_staticSheetNames = "Description, Settings, BatchMaster, BatchDetail, FileOutput"
End Property

' Supported file category that can be replaced during the calculation
' Only Tran.txt and Obligor.txt are supported in this version; to enable more, just add elements to this array
Public Property Get g_sptReplaceFileNames()
    g_sptReplaceFileNames = Array("Tran.txt", "Obligor.txt")
End Property
