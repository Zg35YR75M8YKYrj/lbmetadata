
Imports Microsoft.WindowsAPICodePack
Imports Microsoft.VisualStudio.Shell.Interop
Imports Shell32
Imports System.IO




Module Module1



    Sub Main()
        Dim FileName As String
        FileName = "e:\superbowls.xlsx"
        Dim Properties As Dictionary(Of Integer, KeyValuePair(Of String, String)) = GetFileProperties(FileName)
        For Each FileProperty As KeyValuePair(Of Integer, KeyValuePair(Of String, String)) In Properties
            Console.WriteLine("{0}: {1}", FileProperty.Value.Key, FileProperty.Value.Value)
        Next
        Console.ReadKey()
    End Sub

    Public Function GetFileProperties(ByVal FileName As String) As Dictionary(Of Integer, KeyValuePair(Of String, String))
        Dim Shell As New Shell
        Dim Folder As Folder = Shell.[NameSpace](Path.GetDirectoryName(FileName))
        Dim File As FolderItem = Folder.ParseName(Path.GetFileName(FileName))
        Dim Properties As New Dictionary(Of Integer, KeyValuePair(Of String, String))()
        Dim Index As Integer
        Dim Keys As Integer = Folder.GetDetailsOf(File, 0).Count
        For Index = 0 To Keys - 1
            Dim CurrentKey As String = Folder.GetDetailsOf(Nothing, Index)
            Dim CurrentValue As String = Folder.GetDetailsOf(File, Index)
            If CurrentValue <> "" Then
                Properties.Add(Index, New KeyValuePair(Of String, String)(CurrentKey, CurrentValue))
            End If
        Next
        Return Properties
    End Function

End Module
