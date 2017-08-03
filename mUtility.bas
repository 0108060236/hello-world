Attribute VB_Name = "mUtility"
' @(h) mUtility.BAS ver 1.0 ( 2017.08.01 Numtech)

' @(s)
'
' This module includes utility functions

Option Explicit

Private m_parameterCollection As Collection

'
' Function       : Delete the result worksheets (i.e., ones starting with "Stat")
'
' Return           : A boolean value that indicates whether the operation is sucessful
'
' Argument     : None
'
' Feature          :
'
' Remarks        :

Public Function DeleteResultSheets() As Boolean
    
    DeleteResultSheets = False
    
    Dim ws As Worksheet
    Application.DisplayAlerts = False

    For Each ws In ActiveWorkbook.Sheets
        ' Determine if the sheet is a input sheet
        If InStr(g_staticSheetNames, ws.Name) = 0 Then
            ws.Delete
        End If
    Next ws
    
    Application.DisplayAlerts = True
    DeleteResultSheets = True
    
End Function

'
' Function       : Pop up a diaglog box for entering password
'
' Return           : String to store the password typed in by the users
'
' Argument     : None
'
' Feature          :
'
' Remarks        :

Public Function PasswordInputBox() As String

    Dim password As String
    password = InputBox("Please enter password your NtInsight account to proceed")
    
    If StrPtr(password) = 0 Then
        End
    Else
        PasswordInputBox = password
    End If
    
End Function

'
' Function        : Get the parameter value of a user input, which will be stored as global get only property in mGlobal
'
' Return            : The parameter value of an item in a collection with specified key
'
' Argument      : key - the key for the item
'
' Feature           :
'
' Remarks         :

Public Function GetParamValue(ByVal key As String) As Variant

    ' Parameter will be reloaded, if not yet
    If (m_parameterCollection Is Nothing) Then
        Call LoadParameter
    End If
    
On Error GoTo ErrorHandler

    GetParamValue = m_parameterCollection(key)

    Exit Function
    
ErrorHandler:

    Call logger.OutputError("Parameter is not set. Parameter " & key, Err)
    Call MsgBox("Parameter is not set. Parameter " & key & vbCrLf & "Please re-try after filling the necessary fields", vbCritical)
    Set m_parameterCollection = Nothing
    End

End Function

'
' Function  : Load the parameter values specified in worksheet "Settings"
'
' Return      : None
'
' Argument : None
'
' Feature     :
'
' Remarks   :

Public Function LoadParameter()

    Dim ar As Variant
    Dim i   As Long
    
    Set m_parameterCollection = New Collection

    ar = wsParameter.Range("A1").CurrentRegion
    For i = 2 To UBound(ar) 'First line is the header, starting from the second
        Call m_parameterCollection.Add(ar(i, 2), ar(i, 1))
    Next
    
End Function

'
' Function      : Count number of unique values in a range
'
' Return         : Number of unique values in a range
'
' Argument    : rng - Excel range
'
' Feature        :
'
' Remarks      :

Public Function CountUnique(ByVal rng As Range) As Integer

    Dim dict As New Dictionary
    Set dict = GetUniqueValues(rng)
    
    CountUnique = dict.Count
    
End Function

'
' Function      : Get unique values in a range
'
' Return         : A dictionary that includes the unique values in a range
'
' Argument    : rng - Excel range
'
' Feature       :
'
' Remarks     :

Public Function GetUniqueValues(ByVal rng As Range) As Dictionary

    Dim dict As New Dictionary
    Dim cell As Range
    
    For Each cell In rng.Cells
      If Not dict.Exists(cell.Value) Then
         dict.Add cell.Value, 0
     End If
    Next
    
    Set GetUniqueValues = dict
    
End Function

'
' Function    : Delete a specific worksheet
'
' Return        : None
'
' Argument : sheetName - name of worksheet to delete
'
' Feature      :
'
' Remarks    :

Public Sub DeleteSheet(ByVal sheetName As String)
    
    Dim i   As Integer

    For i = 1 To Sheets.Count
        If sheetName = Sheets(i).Name Then
            Application.DisplayAlerts = False
            Sheets(i).Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next i
    
End Sub

'
' Function       : Create a directory
'
' Return           :
'
' Argument    :
'
' Feature         : If the directory already exists, the function will delete all the files in the directory
'
' Remarks       :

Public Sub CreateDir(ByVal fullPath As String, ByVal filePrefix As String)
    
    If Dir(fullPath, vbDirectory) = "" Then
        MkDir fullPath
    Else
        Call EmptyDir(fullPath, filePrefix)
    End If
    
End Sub

'
' Function   : Delete all the files in a folder directory path
'
' Return       :
'
' Argument  :
'
' Feature     :
'
' Remarks   :

Public Sub EmptyDir(ByVal dirPath As String, ByVal filePrefix As String)
    
    Dim fileName As String
    fileName = Dir(dirPath & "\" & filePrefix & "*.*")
    
    Do While fileName <> ""
      Kill dirPath & "\" & fileName
      fileName = Dir
    Loop
    
End Sub

'
' Function   : Delete a folder directory path
'
' Return       :
'
' Argument :
'
' Feature     :
'
' Remarks   :

Public Sub DeleteDir(ByVal dirPath As String)
    
    On Error Resume Next
    Kill dirPath & "\*.*" ' delete all files in the folder
    RmDir dirPath & "\" ' delete folder
    On Error GoTo 0

End Sub

'
' Function   : Convert tab-delimited text file to csv format
'
' Return       :
'
' Argument :
'
' Feature     :
'
' Remarks   :

Public Sub TxtFileOutput(ByVal filePath As String, ByVal ar As Variant, Optional isTabSeparator As Boolean = True)

    Dim i As Long
    Dim ii As Long
    Dim buf As String

    Open filePath For Output As #1

    For i = 1 To UBound(ar, 1)
        For ii = 1 To UBound(ar, 2)
            If ii = 1 Then
                If Len(ar(i, ii)) = 0 Then
                    buf = ""
                    Exit For
                Else
                    buf = ar(i, ii)
                End If
            Else
                If isTabSeparator Then
                    buf = buf & vbTab & ar(i, ii)
                Else
                    buf = buf & "," & ar(i, ii)
                End If
            End If
        Next
        
        If Len(buf) = 0 Then
            Exit For
        End If
        Print #1, buf
    Next
    
    Close #1

End Sub

'
' Function   : Use regular expression to get the specified text context
'
' Return       : A string of specified text split by space
'
' Argument  :
'
' Feature      :
'
' Remarks   : Default setting of igorning case and global search

Public Function RegularExpression(ByVal stringToSearch As String, ByVal regExPattern As String) As String

    Dim regEx As New RegExp
    Dim matches As Object, match As Object
    Dim outputText As String
    
    With regEx
        .Pattern = regExPattern     ' find pattern
        .IgnoreCase = False
        .Global = True             ' find all the matches
    End With

    Set matches = regEx.Execute(stringToSearch)
    For Each match In matches
        outputText = outputText & " " & match
    Next match
    
    RegularExpression = outputText
    
End Function

'
' Function     : Joins up to 30 collections into one larger collection. (It is not used in the current version)
'
' Return        : Collection that contains all the items from the subcollections passed to this function
'
' Argument : Collections to be merged into larger collection
'                         This parameter is of a ParamArray type, so it is possible to pass any number of collection to be joined (up to 30)
'
' Feature     : All non-collection parameters are ignored
'
' Remarks   : Items in the result collection can be duplicated
'                        Items are added to the result collection without key, even if they are stored

Public Function JoinCollections(ParamArray collections() As Variant) As Collection

    Dim subcollection As Variant
    Dim item As Variant

    'Initialize result variable with empty Collection
    Set JoinCollections = New Collection

    'Iterate through all the given subcollections
    For Each subcollection In collections
        'Check if an item currently processed in this loop is a collection
        If TypeOf subcollection Is Collection Then
            'If [subcollection] is a collection, iterate through all its
            'items and add them to the result Collection [joinCollections]
            For Each item In subcollection
                Call JoinCollections.Add(item)
            Next item
        End If
    Next subcollection
    
End Function
