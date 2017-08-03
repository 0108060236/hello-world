VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBatchDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @(h) clsBatchDetails.cls ver 1.0 ( 2017.08.01 Numtech)
' test
' @(s)
'
' This class stores type-safe collection of clsBatchDetail objects, along with methods belong to the collection

Option Explicit

Private m_batchDetails As Collection

'
' Function     : Get the replaced file names of a single batch
'
' Return        : Dictionary of file names, with file category as keys
'
' Argument   : batchId - the identifier of the batch
'
' Feature      : Both the file name (item) and file category (key) have a postfix of "*.txt" (e.g., batchFileName:="Tran01.txt", key:="Tran.txt")
'
' Remarks    :

Public Function GetBatchFileNames(ByVal batchId As Integer) As Dictionary
    
    Dim item As Object
    Dim batchFileNames As New Dictionary

    For Each item In m_batchDetails
        If item.batchId = batchId Then
            batchFileNames.Add item:=item.batchFileName & ".txt", key:=item.fileCategory & ".txt"
        End If
    Next item
    
    Set GetBatchFileNames = batchFileNames

End Function

'
' Function   : Initialize object, create space for new collection of clsBatchDetail
'
' Return      : None
'
' Argument : None
'
' Feature      :
'
' Remarks    :
Private Sub Class_Initialize()

    Set m_batchDetails = New Collection

End Sub

'
' Function   : Add function to add element to the collection
'
' Return      : The added clsBatchDetail object
'
' Argument : batchId - the identifier of the batch, field of clsBatchDetail
'                        fileCategory - type of the file to be replaced, field of clsBatchDetail
'                        batchFileName - name of the replaced file, field of clsBatchDetail
'                        key - key of the record
'
' Feature      :
'
' Remarks    :
Public Function Add(ByVal batchId As Integer, ByVal fileCategory As String, ByVal batchFileName As String, Optional key As String) As clsBatchDetail
    
    Dim objNewMember As clsBatchDetail
    Set objNewMember = New clsBatchDetail
    
    ' Assign values
    With objNewMember
        .batchId = batchId
        .fileCategory = fileCategory
        .batchFileName = batchFileName
    End With

    If Len(key) = 0 Then
        m_batchDetails.Add objNewMember
    Else
        m_batchDetails.Add objNewMember, key
    End If
    
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

'
' Function   : Function to access item in the typed collection
'
' Return      : The clsBatchDetail object with the specified index key
'
' Argument : vntIndex - the index to access the item, and it can be either index or key
'
' Feature      :
'
' Remarks    :
Public Property Get item(vntIndexKey As Variant) As clsBatchDetail

    Err.Clear
    On Error GoTo GetItemError
    Set item = m_batchDetails(vntIndexKey)
    Exit Property
    
GetItemError:
    Debug.Print Err.Source & " : Error Was found in Item(vntIndexKey as Variant) as clsbatchDetails . " & vbCrLf & Err.Description & " : " & Err.Number
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

    Count = m_batchDetails.Count
    
End Property
