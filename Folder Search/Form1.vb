Public Class Form1

    Dim AllFolders() As Object
    Dim index As Integer

    Dim ol As New Outlook.Application
    Dim myNameSpace As Outlook.NameSpace



    Private Sub ListBox1_MouseClick(sender As Object, e As Windows.Forms.MouseEventArgs) Handles ListBox1.MouseClick
        Dim Folders() As String

        Dim FolderPath As String
        Dim ListBoxPath As String
        ListBoxPath = ListBox1.SelectedItem

        FolderPath = Microsoft.VisualBasic.Strings.Right(ListBox1.SelectedItem, Len(ListBox1.SelectedItem) - 2)
        Folders = Split(FolderPath, "\", , vbTextCompare)
        Dim antallMapper As Integer
        antallMapper = 0

        antallMapper = UBound(Folders)


        Select Case antallMapper

            Case 0
                ol.Application.Session.Folders.Item(Folders(0)).Display()

            Case 1
                ol.Application.Session.Folders.Item(Folders(0)).Folders.Item(Folders(1)).Display()

            Case 2
                ol.Application.Session.Folders.Item(Folders(0)).Folders.Item(Folders(1)).Folders.Item(Folders(2)).Display()

            Case 3
                ol.Application.Session.Folders.Item(Folders(0)).Folders.Item(Folders(1)).Folders.Item(Folders(2)) _
                .Folders.Item(Folders(3)).Display()

            Case 4
                ol.Application.Session.Folders.Item(Folders(0)).Folders.Item(Folders(1)).Folders.Item(Folders(2)) _
                .Folders.Item(Folders(3)).Folders.Item(Folders(4)).Display()

            Case 5
                ol.Application.Session.Folders.Item(Folders(0)).Folders.Item(Folders(1)).Folders.Item(Folders(2)) _
                .Folders.Item(Folders(3)).Folders.Item(Folders(4)).Folders.Item(Folders(5)).Display()

            Case 6
                ol.Application.Session.Folders.Item(Folders(0)).Folders.Item(Folders(1)).Folders.Item(Folders(2)) _
                .Folders.Item(Folders(3)).Folders.Item(Folders(4)).Folders.Item(Folders(5)).Folders.Item(Folders(6)).Display()

            Case 7
                ol.Application.Session.Folders.Item(Folders(0)).Folders.Item(Folders(1)).Folders.Item(Folders(2)) _
                .Folders.Item(Folders(3)).Folders.Item(Folders(4)).Folders.Item(Folders(5)).Folders.Item(Folders(6)) _
                .Folders.Item(Folders(7)).Display()

            Case 8
                ol.Application.Session.Folders.Item(Folders(0)).Folders.Item(Folders(1)).Folders.Item(Folders(2)) _
                .Folders.Item(Folders(3)).Folders.Item(Folders(4)).Folders.Item(Folders(5)).Folders.Item(Folders(6)) _
                .Folders.Item(Folders(7)).Folders.Item(Folders(8)).Display()

            Case 9
                ol.Application.Session.Folders.Item(Folders(0)).Folders.Item(Folders(1)).Folders.Item(Folders(2)) _
                .Folders.Item(Folders(3)).Folders.Item(Folders(4)).Folders.Item(Folders(5)).Folders.Item(Folders(6)) _
                .Folders.Item(Folders(7)).Folders.Item(Folders(8)).Folders.Item(Folders(9)).Display()

            Case Is >= 10
                MsgBox("To many subfolders", vbExclamation, "Folder Search")


        End Select

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListBox1.Items.Clear()
        myNameSpace = ol.GetNamespace("MAPI")


        Dim myOutlookFolder As Outlook.Folder
        index = 0
        ReDim Preserve AllFolders(index)


        For Each myOutlookFolder In myNameSpace.Folders
            AllFolders(index) = myOutlookFolder.FolderPath
            index = index + 1
            ReDim Preserve AllFolders(index)
            Call SearchSubFolders(myOutlookFolder)

        Next

        Call SearchResults()
    End Sub

    Private Sub SearchSubFolders(Foldername As Outlook.Folder)

        Dim SubFolder As Outlook.Folder

        For Each SubFolder In Foldername.Folders
            AllFolders(index) = SubFolder.FolderPath
            index = index + 1
            ReDim Preserve AllFolders(index)
            Call SearchSubFolders(SubFolder)
        Next

    End Sub
    Private Sub SearchResults()
        Dim str As Object
        Dim searchindex As Integer
        For Each str In AllFolders
            Dim startIndex As Integer
            startIndex = InStrRev(str, "\")
            searchindex = InStr(startIndex + 1, str, My.Settings.strSearchString, vbTextCompare)

            If searchindex > 0 Then
                ListBox1.Items.Add(str)
            End If

        Next

    End Sub

    

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

        Me.Close()

    End Sub

    
End Class