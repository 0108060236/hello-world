VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @(h) clsLog.BAS ver 1.0 ( 2017.08.01 Numtech)

' @(s)
'
' This class controls the log information

Option Explicit

' Log file type extension
Private Const LOG_FILE_EXTENSION As String = ".log"

'FileSystemObject
Private objFs As Object

' IO mode attribute for TextStreamObject
Private Enum IOMode
    ForReading = 0      ' Read only mode (default value)
    ForWriting = 1      ' Write mode
    ForAppending = 8    ' Append mode
End Enum

'
' Function     : Initialize IO object
'
' Return        : None
'
' Argument   : None
'
' Feature        :
'
' Remarks      :

Private Sub Class_Initialize()

    Set objFs = CreateObject("Scripting.FileSystemObject")
    
End Sub

'
' Function      : Terminate IO object
'
' Return          : None
'
' Argument  : None
'
' Feature       :
'
' Remarks     :

Private Sub Class_Terminate()

    Set objFs = Nothing
    
End Sub

'
' Function   : Output log [Information]
'
' Return       : None
'
' Argument : msg - message
'
' Feature      :
'
' Remarks    :

Public Sub OutputInfo(ByVal msg As String)
    
    Call OutputLog("Information", msg)
    
End Sub

'
' Function     : Output log [Warning]
'
' Return          : None
'
' Argument  : msg - message
'
' Feature       :
'
' Remarks     :

Public Sub OutputWarn(ByVal msg As String)
    
    Call OutputLog("Warning", msg)
    
End Sub

'
' Function     : Output log [Error]
'
' Return         : None
'
' Argument  : msg - message
'                         objErr - error object
'
' Feature       :
'
' Remarks     :

Public Sub OutputError(ByVal msg As String, Optional ByVal objErr As ErrObject = Nothing)
    
    Dim strMsg As String
    
    strMsg = msg
    
    If Not (objErr Is Nothing) Then
        strMsg = vbTab & strMsg & ":" & "Err.Number:[" & objErr.Number & "]," & "Err.Description:[" & objErr.Description & "]:"
    End If
    
    Call OutputLog("Error", strMsg)

End Sub

'
' Function        : Output log.
'
' Return            : None
'
' Argument      : logType - log information type
'                              msg - message
'
' Feature           :
'
' Remarks         :

Private Sub OutputLog(ByVal logType As String, ByVal msg As String)
    
    Dim objTs As Object     ' TextStreamObject
    Dim i As Long
    
    On Error GoTo LBL_ERR:
    
    Set objTs = objFs.OpenTextFile( _
        fileName:=objFs.BuildPath(ThisWorkbook.Path, ThisWorkbook.Name & LOG_FILE_EXTENSION), _
        IOMode:=IOMode.ForAppending, _
        Create:=True)

    objTs.WriteLine Format(Now, "yyyy/mm/dd hh:mm:ss") & vbTab & logType & vbTab & msg

    objTs.Close
    Set objTs = Nothing

   Exit Sub
LBL_ERR:
   
    If Not (objTs Is Nothing) Then
        objTs.Close
        Set objTs = Nothing
    End If

    Err.Raise Number:=Err.Number, Description:="Log output failed." & Err.Description

End Sub

