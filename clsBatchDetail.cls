VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBatchDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @(h) clsBatchDetail.cls ver 1.0 ( 2017.08.01 Numtech)

' @(s)
'
' This class defines data structure for data contained in worksheet "BatchDetail"

Option Explicit

Public batchId As Integer   ' A non-decreasing integer that starts from 1
Public fileCategory As String   ' Standard file name, ".txt" is not included
Public batchFileName As String      ' User defined file names
