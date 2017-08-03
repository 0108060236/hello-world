VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNicmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @(h) clsNicmd.BAS ver 1.0 ( 2017.08.01 Numtech)

' @(s)
'
' This class performs basic functions of nicmd

Option Explicit

Private m_nicmdObject As Object
Private m_nicmdPrompt As Variant
Private m_output As String

'
' Function   : Initialize object
'
' Return      : None
'
' Argument : None
'
' Feature      :
'
' Remarks    :

Private Sub Class_Initialize()

    ' Do nothing

End Sub

'
' Function  : Terminate object
'
' Return      : None
'
' Argument : None
'
' Feature      :
'
' Remarks    :

Private Sub Class_Terminate()

    Quit

End Sub

'
' Function  : It reads relevant setting information of nicmd, and connect with NIC
'
' Return      : Boolean value that indicates whether the connection is successful
'
' Argument : None
'
' Feature      :
'
' Remarks   :

Public Function NicmdConfig() As Boolean

    NicmdConfig = False

    ' Connection information
    ' For security reason, password will be enterred manually by the user with a pop up window
    Dim password As String
    password = PasswordInputBox
    
    If Not Connect(g_pathNicmd, g_hostName, g_portNo, g_userName, password) Then
        Call logger.OutputError("Connection error(host name, port number, user name, or password is incorrect)", Err)
        Exit Function
    End If
    
    NicmdConfig = True

End Function

'
' Function            : Connect to NtInsight for Credit Risk
'
' Return                : Boolean value that indicates whether the connection is successful
'
' Argument          : nicmdFullPath - full file path of nicmd.exe
'                                  hostName - host name
'                                  portNo - port number
'                                  userName - user name
'                                  password - password
'
' Feature              :
'
' Remarks            :

Public Function Connect(ByVal nicmdFullPath As String, ByVal hostName As String, ByVal portNo As Long, ByVal userName As String, ByVal password As String) As Boolean
    
    Connect = False
    
    Dim nicmdExecText As String

    nicmdExecText = """" & nicmdFullPath & """ -h " & hostName & " -P " & portNo & " -u " & userName & " -p " & IIf(Len(password) = 0, "", password)
    m_nicmdPrompt = Array("n", "i", "c", "m", "d", ">", " ")
    
    ' Cannot print nicmdExecText directly, as password should not be revealed
    Call logger.OutputInfo("nicmd: Connecting to server...")
    
    With CreateObject("Wscript.Shell")
        Set m_nicmdObject = .exec(nicmdExecText)
    End With
    
    Call logger.OutputInfo("nicmd: Connection successful.")
    
    ParseResponse

    Connect = IsConnected()
    
End Function

'
' Function  : Pass commands to NtInsight for Credit Risk
'
' Return      :  Boolean value that indicates whether the connection is successful
'
' Argument  : nicmdCommand - command string
'
' Feature     :
'
' Remarks   :

Private Function ExecuteCommand(ByVal nicmdCommand As String) As Boolean

    If Not IsConnected() Then
        ExecuteCommand = False
        Exit Function
    End If
    
    ' Write to log about the commands
    Call logger.OutputInfo("Executing nicmd comand: " & nicmdCommand)
    
    m_nicmdObject.StdIn.WriteLine (nicmdCommand)
    
    m_output = ParseResponse()
    ExecuteCommand = Not IsError()

End Function

'
' Function              : Check whether the execution results in an error
'
' Return                  : Boolean value that indicates whether any error occurs
'
' Argument            : None
'
' Feature                 :
'
' Remarks               :

Private Function IsError() As Boolean

    IsError = (Len(m_output) >= 6) And (Left(m_output, 6) = "ERROR:")

End Function

'
' Function              : Check whether nicmd is connected
'
' Return                  : Boolean value that indicates whether nicmd is connect
'
' Argument            : None
'
' Feature                :
'
' Remarks              :

Public Function IsConnected() As Boolean

    IsConnected = m_nicmdObject.Status = 0

End Function

'
' Function             : Quit nicmd
'
' Return                 : Boolean value that indicates whether the command execution is successful
'
' Argument          : None
'
' Feature               :
'
' Remarks             :

Public Function Quit() As Boolean

    If Not m_nicmdObject Is Nothing Then
        If IsConnected() Then
            m_nicmdObject.StdIn.WriteLine "Quit"
        End If
        Set m_nicmdObject = Nothing
    End If
    
    Quit = True
    
End Function

'
' Function    : Get file names from a workspace
'
' Return        : A collection that splits file names into xml file, input files (xml and txt input), and output files
'
' Argument : workspace - destination workspace to get file names from
'
' Feature      :
'
' Remarks    : Every item of the collection is a string of file names split by space

Public Function GetFileNames(ByVal workspace As String) As Collection
    
    Dim stringToSearch As String
    Dim regExPattern As String
    Dim fileNames As New Collection
    
    ' Get information of workspace's files
    Dirws workspace
    ' Parse the file information to file names split by space
    stringToSearch = RegularExpression(m_output, "\.*[^\\]*\.txt|\.*[^\\]*\.xml")

    ' Get xml file names (e.g., Cb.xml, CbSchema.xml), can be retrieved by key "XML"
    regExPattern = ".*?xml"
    fileNames.Add item:=RegularExpression(stringToSearch, regExPattern), key:="XML"

    ' Get input file names, including both xml and txt input, can be retrieved by key "INPUT"
    regExPattern = " (.*?xml)| (?!Stat).*?txt"
    fileNames.Add item:=RegularExpression(stringToSearch, regExPattern), key:="INPUT"
    
    ' Get output file names, can be retrieved by key "STAT"
    regExPattern = "Stat.*?txt"
    fileNames.Add item:=RegularExpression(stringToSearch, regExPattern), key:="STAT"

    Set GetFileNames = fileNames
    
End Function

'
' Function    : Get file information of a workspace / List of workspace's files
'
' Return        : Boolean value that indicates whether the command execution is successful
'
' Argument : workspace - destination workspace to get file names from
'
' Feature     :
'
' Remarks   :

Private Function Dirws(ByVal workspace As String) As Boolean

    Dim nicmdCommands As String
    
    nicmdCommands = "dirws " & workspace & ""
    Dirws = ExecuteCommand(nicmdCommands)

End Function

'
' Function   : Upload text files to NtInsight for Credit Risk
'
' Return      : Boolean value that indicates whether the command execution is successful
'
' Argument : workspace - destination workspace to upload to
'                        localFolderPath - source file folder which contains files to upload
'                        fileName - name of text file to upload
'
' Feature     :
'
' Remarks   :

Public Function Upload(ByVal workspace As String, ByVal localFolderPath As String, ByVal fileName As String) As Boolean

    Dim nicmdCommands As String
    
    nicmdCommands = "upload " & workspace & " resultonly """ & localFolderPath & """ " & fileName
    Upload = ExecuteCommand(nicmdCommands)

End Function

'
' Function    : Run simulation with NtInsight for Credit Risk
'
' Return        : Boolean value that indicates whether the command execution is successful
'
' Argument : workspace - workspace to run simulation for
'
' Feature      :
'
' Remarks    :

Public Function Resimulate(ByVal workspace As String) As Boolean

    Dim nicmdCommands As String
    
    nicmdCommands = "resimulate " & workspace
    Resimulate = ExecuteCommand(nicmdCommands)

End Function

'
' Function     : Download text file from NtInsight for Credit Risk workspace
'
' Return         : Boolean value that indicates whether the command execution is successful
'
' Argument  : workspace - source of workspace
'                         localFolderPath - destination folder path
'                         fileName - name of text file to download
'
' Feature      :
'
' Remarks    :

Public Function Download(ByVal workspace As String, ByVal localFolderPath As String, ByVal fileName As String) As Boolean

    Dim nicmdCommands As String
    
    nicmdCommands = "download " & workspace & " """ & localFolderPath & """ " & fileName
    Download = ExecuteCommand(nicmdCommands)

End Function

'
' Function      : Split command execution result to elements of an array
'
' Return          : Array of command execuation result
'
' Argument   : None
'
' Feature        :
'
' Remarks      :

Public Function GetOutputStringAsArray() As Variant

    Dim arr As Variant
    
    arr = Split(m_output, vbCrLf)
    GetOutputStringAsArray = arr

End Function

'
' Function        : Parse command execution result as string
'
' Return            : Command execution result string
'
' Argument      : None
'
' Feature           :
'
' Remarks         :

Public Function ParseResponse() As String

    Dim output As String
    Dim header As String
    Dim c As String
    Dim i As Long
    Dim isPrompt As Boolean
        
    output = ""
    Do While True
        If Not IsConnected() Then
            output = "ERROR:ParseResponse"
            Exit Do
        End If
        
         header = ""
        isPrompt = True
        For i = LBound(m_nicmdPrompt) To UBound(m_nicmdPrompt)
            If Not IsConnected() Then
                Exit For
            End If
            
            c = m_nicmdObject.StdOut.Read(1)
            header = header & c
            If m_nicmdPrompt(i) <> c Then
                isPrompt = False
                Exit For
            End If
        Next i
        
        If isPrompt Then
            Exit Do
        End If
        
        output = output & header & m_nicmdObject.StdOut.ReadLine() & vbCrLf
    Loop

    ParseResponse = output

End Function
