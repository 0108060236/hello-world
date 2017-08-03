VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalcConfigs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @(h) clsCalcConfigs.cls ver 1.0 ( 2017.08.01 Numtech)

' @(s)
'
' This class stores type-safe collection of clsCalcConfig objects

Option Explicit

Private m_calcConfigs As Collection

'
' Function   : Initialize object, create space for new collection of clsCalcConfig
'
' Return      : None
'
' Argument : None
'
' Feature      :
'
' Remarks    :

Private Sub Class_Initialize()

    Set m_calcConfigs = New Collection

End Sub

'
' Function   : Add function to add element to the collection
'
' Return      : The added clsCalcConfig object
'
' Argument : batchId - the identifier of the batch, field of clsCalcConfig
'                        batchName - name of the batch, field of clsCalcConfig
'                        noOfSystemIterations - calculation setting node to be modified, field of clsCalcConfig
'                        key - key of the record
'
' Feature      :
'
' Remarks    :

Public Function Add(ByVal batchId As Integer, ByVal batchName As String, ByVal noOfSystemIterations As Integer, Optional key As String) As clsCalcConfig
    
    Dim objNewMember As clsCalcConfig
    Set objNewMember = New clsCalcConfig
    
    ' Assign values
    With objNewMember
        .batchId = batchId
        .batchName = batchName
        .noOfSystemIterations = noOfSystemIterations
    End With

    If Len(key) = 0 Then
        m_calcConfigs.Add objNewMember
    Else
        m_calcConfigs.Add objNewMember, key
    End If
    
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

'
' Function   : Function to access item in the typed collection
'
' Return      : The clsCalcConfig object with the specified index key
'
' Argument : vntIndex - the index to access the item, and it can be either index or key
'
' Feature      :
'
' Remarks    :

Public Property Get item(vntIndexKey As Variant) As clsCalcConfig

    Err.Clear
    On Error GoTo GetItemError
    Set item = m_calcConfigs(vntIndexKey)
    Exit Property
    
GetItemError:
    Debug.Print Err.Source & " : Error Was found in Item(vntIndexKey as Variant) as clsCalcConfigs . " & vbCrLf & Err.Description & " : " & Err.Number
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

    Count = m_calcConfigs.Count
    
End Property


