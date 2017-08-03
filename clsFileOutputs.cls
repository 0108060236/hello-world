VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileOutputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @(h) clsFileOutputs.cls ver 1.0 ( 2017.08.01 Numtech)

' @(s)
'
' This class stores type-safe collection of clsFileOutput objects

Option Explicit

Private m_fileOutputs As Collection

'
' Function   : Initialize object, create space for new collection of fileOutput
'
' Return      : None
'
' Argument : None
'
' Feature      :
'
' Remarks    :

Private Sub Class_Initialize()

    Set m_fileOutputs = New Collection

End Sub

'
' Function   : Add function to add element to the collection
'
' Return      : The added clsFileOutput object
'
' Argument : fileCategory - standard file type, field of clsFileOutput
'                        fileheader - header of output txt file, field of clsFileOutput
'                        fileNames - names of files for the file category, field of clsFileOutput
'                        key - key of the record
'
' Feature      :
'
' Remarks    :

Public Function Add(ByVal fileCategory As String, ByVal fileHeader As String, ByVal attributeName As String, ByVal fileNames As Collection, Optional key As String) As clsFileOutput
    
    Dim objNewMember As clsFileOutput
    Set objNewMember = New clsFileOutput
    
    ' Assign values
    With objNewMember
        .fileCategory = fileCategory
        .fileHeader = fileHeader
        .attributeName = attributeName
        Set .fileNames = fileNames
    End With

    If Len(key) = 0 Then
        m_fileOutputs.Add objNewMember
    Else
        m_fileOutputs.Add objNewMember, key
    End If
    
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

'
' Function   : Function to access item in the typed collection
'
' Return      : The clsFileOutput object with the specified index key
'
' Argument : vntIndex - the index to access the item, and it can be either index or key
'
' Feature      :
'
' Remarks    :

Public Property Get item(vntIndexKey As Variant) As clsFileOutput

    Err.Clear
    On Error GoTo GetItemError
    Set item = m_fileOutputs(vntIndexKey)
    Exit Property
    
GetItemError:
    Debug.Print Err.Source & " : Error Was found in Item(vntIndexKey as Variant) as clsFileOutputs . " & vbCrLf & Err.Description & " : " & Err.Number
    Set item = Nothing
    
End Property

'
' Function   : Count the number of objects in the typed collection
'
' Return       : An integer indicates the number of elements in the collection
'
' Argument :
'
' Feature      :
'
' Remarks    :

Public Property Get Count() As Integer

    Count = m_fileOutputs.Count
    
End Property

