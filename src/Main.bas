Attribute VB_Name = "Main"
Option Explicit

Const colModelNames As Integer = 0
Const colConfNames As Integer = 1
Const targetDirName As String = ".SaveAsPart"

Dim swApp As Object
Dim confCollection As Dictionary  'string: ConfigurationItem
Dim gFSO As FileSystemObject
Dim openedDocs As Collection 'ModelDoc2
Dim currentDoc As ModelDoc2

Sub Main()
    Set swApp = Application.SldWorks
    Set currentDoc = swApp.ActiveDoc
    If currentDoc Is Nothing Then Exit Sub
    If currentDoc.GetType <> swDocASSEMBLY Then
        MsgBox "Макрос работает только в сборках.", vbCritical
        Exit Sub
    End If
    
    Init
    CollectItems
    MainForm.txtTargetDir.Text = gFSO.GetParentFolderName(confCollection(0).Model.GetPathName) + "\" + targetDirName
    MainForm.Show
End Sub

Function Init() 'mask for button
    Dim i As Variant
    
    Set gFSO = New FileSystemObject
    Set confCollection = New Dictionary
    Set openedDocs = New Collection
    For Each i In swApp.Frame.ModelWindows
        openedDocs.Add i.ModelDoc
    Next
End Function

Function CollectItems() 'mask for button
    Dim i As Integer
    Dim doc As ModelDoc2
    Dim path As String
    Dim errors As swFileLoadError_e
    Dim warnings As swFileLoadWarning_e
    Dim depends As Variant

    AddConfigurationItem currentDoc
    depends = currentDoc.GetDependencies2(True, True, False)
    For i = LBound(depends) To UBound(depends) Step 2
        path = depends(i + 1)
        If UCase(Right(path, 6)) = "SLDASM" Then
            Set doc = swApp.OpenDoc6(path, swDocASSEMBLY, _
                swOpenDocOptions_OverrideDefaultLoadLightweight + swOpenDocOptions_LoadLightweight + swOpenDocOptions_ReadOnly, _
                "", errors, warnings)
            If Not doc Is Nothing Then
                AddConfigurationItem doc
            End If
        End If
    Next
End Function

Sub AddConfigurationItem(doc As ModelDoc2)
    Dim rowNumber As Integer
    Dim BaseName As String
    Dim i As Integer
    Dim confNames() As String
    
    BaseName = gFSO.GetBaseName(doc.GetPathName)
    confNames = doc.GetConfigurationNames
    For i = 0 To doc.GetConfigurationCount - 1
        rowNumber = MainForm.lstConfs.ListCount 'must be first!
        MainForm.lstConfs.AddItem
        MainForm.lstConfs.List(rowNumber, colModelNames) = BaseName
        MainForm.lstConfs.List(rowNumber, colConfNames) = confNames(i)
        confCollection.Add rowNumber, New ConfigurationItem
        With confCollection(rowNumber)
            Set .Model = doc
            .BaseName = BaseName
            .ConfName = confNames(i)
        End With
    Next
End Sub

Function Run() 'mask for button
    Dim i As Integer
    Dim item As ConfigurationItem
    Dim targetDir As String
    Dim newName As String
    Dim errors As swFileSaveError_e
    Dim warnings As swFileSaveWarning_e
    Dim errors2 As swActivateDocError_e
    Dim openedDoc As Variant
    Dim currentConf As String

    targetDir = MainForm.txtTargetDir.Text
    If Not gFSO.FolderExists(targetDir) Then
        gFSO.CreateFolder targetDir
    End If
    'TODO: check if cannot create dir
    
    currentConf = currentDoc.ConfigurationManager.ActiveConfiguration.Name
    For i = MainForm.lstConfs.ListCount - 1 To 0 Step -1
        If MainForm.lstConfs.Selected(i) Then
            Set item = confCollection(i)
            newName = targetDir + "\" + item.BaseName + "--" + item.ConfName + ".SLDPRT"
            swApp.ActivateDoc3 item.Model.GetPathName, False, swDontRebuildActiveDoc, errors2
            item.Model.ShowConfiguration2 item.ConfName
            item.Model.Extension.SaveAs newName, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, errors, warnings
            
            For Each openedDoc In openedDocs
                If item.Model Is openedDoc Then GoTo NextItem
            Next
            swApp.CloseDoc item.Model.GetPathName
NextItem:
        End If
    Next
    currentDoc.ShowConfiguration2 currentConf
    Shell "explorer """ & targetDir & "", vbNormalFocus
End Function

Function RunAndExit() 'mask for button
    Run
    ExitApp
End Function

Function ExitApp() 'mask for button
    Unload MainForm
    End
End Function

Function SetSelectionAllRows() 'mask for button
    Dim i As Integer
    
    For i = 0 To MainForm.lstConfs.ListCount - 1
        MainForm.lstConfs.Selected(i) = True
    Next
End Function

Function InvertSelectionAllRows() 'mask for button
    Dim i As Integer
    
    For i = 0 To MainForm.lstConfs.ListCount - 1
        MainForm.lstConfs.Selected(i) = Not MainForm.lstConfs.Selected(i)
    Next
End Function
