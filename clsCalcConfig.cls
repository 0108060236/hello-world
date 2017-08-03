VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalcConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @(h) clsCalcConfig.cls ver 1.0 ( 2017.08.01 Numtech)

' @(s)
'
' This class handles tasks for calculation setting file (e.g., Cb.xml)

Option Explicit

Public batchId As Integer
Public batchName As String
Public noOfSystemIterations As Integer

Public Enum productType
    Cb = 1
End Enum

'
' Function     : Overwrite calculation setting files based on user input
'
' Return         : Boolean value that indicates whether the calculation setting is successfully handled
'
' Argument  : None
'
' Feature       :
'
' Remarks     :

Public Function ConfigCalcSetting() As Boolean

    ConfigCalcSetting = False

    ' Load, parse, replace, and save for number of system iterations for Cb
    XmlHelper g_baseWsDataDirPath, "NoOfSystemIterations", noOfSystemIterations, g_baseWsDataDirPath, Cb
    
    ConfigCalcSetting = True
    
End Function

'
' Function     : Load and parse the xml file, replace the tag value and save to target directory
'
' Return         : Boolean value that indicates whether loading, parsing, modification and saving process is successful
'
' Argument  : filePath - folder directory for the calculation setting file
'                         tagName - tag name to get element from
'                         replaceValue - value to replace for the target element
'                         targetDir - directory to save the modified document
'                         productType - product type for generating calculation file name accordingly
'
' Feature       :
'
' Remarks     :

Private Function XmlHelper(ByVal folderPath As String, ByVal tagName As String, ByVal replaceValue As String, ByVal targetDir As String, ByVal productType As productType) As Boolean

    XmlHelper = False
    
    Dim xml As New MSXML2.DOMDocument
    Dim fileName As String
    
    ' Generate xml file name based on product type
    ' Only Cb.xml (NIC) is supported in this version
    Select Case productType
        Case Cb
            fileName = "Cb.xml"
    End Select
    
    xml.Load folderPath & fileName    ' Load the dom document in the directory as domdocument object
    xml.getElementsByTagName(tagName).item(0).Text = replaceValue   ' replace the value of a specified tag
    xml.Save targetDir & fileName     ' Save the modified file to the target directory
    
    XmlHelper = True
    
End Function
