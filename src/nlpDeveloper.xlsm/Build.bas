Attribute VB_Name = "Build"
'''
'這個做法跟原先的 VBADeveloper 不同，原先的 VBADeveloper 做法，請參考 src/vbaDeveloper.xlam 下的 Build.bas。
'理由很簡單，因為我們把 poc_tool8.xlsm 當作 template，需要機動性的使用，而原先的作法是 Addins，會綁絕對路徑:
'1. 如果你是一個新的專案，請你打開一個新的 Excel，例如 A.xlsx
'2. 打開 A.xlsx， VB Scriptor (Alt+F11) 然後 檔案 (Menu) > 匯入檔案 (Import) (Ctrl+M)，將這個 Build.bas 檔案匯入。
'3. 打開 VB Script Editor > Tools references (工具/設定引用項目) 檢查你是否將以下兩個方塊打勾:
'   * Microsoft Visual Basic for Applications Extensibility 5.x
'   * Microsoft Scripting Runtime
'4. 將Project Name 改成 NLPDeveloper (名字無法改變，除非你改程式不然會有Bug)，(Ctrl+R --> F4) 屬性--> Name。
'5. 關閉 VB Script Editor 後，在檔案 > 選項 > 信任中心 > 信任中心設定 > 巨集設定，將 啟用所有巨集 以及 信任存取 VBA 專案物件模型 都勾選起來。
'6. 將 A.xlsx 另存為 xlsm 檔案。 記住不是 xlam 檔。
'7. 開啟 A.xlsm， 打開 VB Script Editor，在 NLPDeveloper 專案下將以下檔案用手動的方式匯入檔案。
'   記住這裏第一次得要手動匯入，因為這些檔案之後都不會再匯出匯入，也就是要改就打開這個 A.xlsm 改。
'   以下這些檔案都可以從 src/nlpDeveloper.xlam 下找到:
'   * CustomActions.cls
'   * ErrorHandling.bas
'   * EventListener.cls
'   * Formatter.bas
'   * Menu.bas
'   * MyCustomActions.cls
'   * NamedRanges.bas
'   * Test.bas
'   * XMLexporter.bas
'8. 在 NLPDeveloper 專案下的 ThisWookbook.cls ，新增以下兩個 function，用來使用 增益集 > NLPDeveloper > Import / Export:
'   Private Sub Workbook_Open()
'       menu.createMenu
'   End Sub
'   Private Sub Workbook_BeforeClose(Cancel As Boolean)
'       menu.deleteMenu
'   End Sub
'9. 關閉 A.xlsm，之後就可以把 A.xlsm 拿來當作 template 使用，不受到 Addins 被需綁絕對路徑的限制，且 Export/Import 也不會把NLPDeveloper 的
'   Version Control Files 匯出匯入。
'10.做完 template 後，第一次想要匯入你自己的檔案，將 A.xlsm 更名成為你的專案名，並與你的VBA 檔案的位置如下放置。
'   假設專案名為 poc_tool8:
'   poc_tool8.xlsm
'     |_ src
'         |_ poc_tool8.xlsm
'             |_ a.bas
'             |_ b.bas
'             |_ c.cls ...
'   放置完畢後，打開 poc_tool8.xlsm > VB Script Editor > Build.bas，然後把你的游標移到 testImport() 上，
'   然後 執行 > 執行 Sub 或 UserForm (F5)，就可以把 src/poc_tool8.xlsm 下你的 VBA code 匯入。 同理，testExport() 亦然。
'11.當然這時你要從 增益集 > NLPDeveloper > Import 也行。
'12.特別提醒，所有以上匯入的檔案，請把 Option Explicit 這行以上(不含此行)的東西 都拿掉，包含 Attribute 或其他宣告。
'13.最後，VBA匯出 ok 後，把修改好的專案 template: poc_tool8.xlsm，放到 templates 目錄下，方便 pyinstaller 包板使用。
'''

Option Explicit

Private Const IMPORT_DELAY As String = "00:00:03"

'We need to make these variables public such that they can be given as arguments to application.ontime()
Public componentsToImport As Dictionary 'Key = componentName, Value = componentFilePath
Public sheetsToImport As Dictionary 'Key = componentName, Value = File object
Public vbaProjectToImport As VBProject

Public Sub testImport()
    Dim proj_name As String
    proj_name = "NLPDeveloper"

    Dim VBAProject As Object
    Set VBAProject = Application.VBE.VBProjects(proj_name)
    Build.importVbaCode VBAProject
End Sub

Public Sub testExport()
    Dim proj_name As String
    proj_name = "NLPDeveloper"

    Dim VBAProject As Object
    Set VBAProject = Application.VBE.VBProjects(proj_name)
    Build.exportVbaCode VBAProject
End Sub

' Returns the directory where code is exported to or imported from.
' When createIfNotExists:=True, the directory will be created if it does not exist yet.
' This is desired when we get the directory for exporting.
' When createIfNotExists:=False and the directory does not exist, an empty String is returned.
' This is desired when we get the directory for importing.
'
' Directory names always end with a '\', unless an empty string is returned.
' Usually called with: fullWorkbookPath = wb.FullName or fullWorkbookPath = vbProject.fileName
' if the workbook is new and has never been saved,
' vbProject.fileName will throw an error while wb.FullName will return a name without slashes.
Public Function getSourceDir(fullWorkbookPath As String, createIfNotExists As Boolean) As String
    ' First check if the fullWorkbookPath contains a \.
    If Not InStr(fullWorkbookPath, "\") > 0 Then
        'In this case it is a new workbook, we skip it
        Exit Function
    End If

    Dim FSO As New Scripting.FileSystemObject
    Dim projDir As String
    projDir = FSO.GetParentFolderName(fullWorkbookPath) & "\"
    Dim srcDir As String
    srcDir = projDir & "src\"
    Dim exportDir As String
    exportDir = srcDir & FSO.GetFileName(fullWorkbookPath) & "\"

    If createIfNotExists Then
        If Not FSO.FolderExists(srcDir) Then
            FSO.CreateFolder srcDir
            Debug.Print "Created Folder " & srcDir
        End If
        If Not FSO.FolderExists(exportDir) Then
            FSO.CreateFolder exportDir
            Debug.Print "Created Folder " & exportDir
        End If
    Else
        If Not FSO.FolderExists(exportDir) Then
            Debug.Print "Folder does not exist: " & exportDir
            exportDir = ""
        End If
    End If
    getSourceDir = exportDir
End Function

' Usually called after the given workbook is saved
Public Sub exportVbaCode(VBAProject As VBProject)
    Dim vbProjectFileName As String
    On Error Resume Next
    'this can throw if the workbook has never been saved.
    vbProjectFileName = VBAProject.fileName
    On Error GoTo 0
    If vbProjectFileName = "" Then
        'In this case it is a new workbook, we skip it
        Debug.Print "No file name for project " & VBAProject.name & ", skipping"
        Exit Sub
    End If

    Dim export_path As String
    export_path = getSourceDir(vbProjectFileName, createIfNotExists:=True)

    Debug.Print "exporting to " & export_path
    'export all components
    Dim component As VBComponent
    For Each component In VBAProject.VBComponents
        'lblStatus.Caption = "Exporting " & proj_name & "::" & component.Name
        '''
        ' Removed condition "If hasCodeToExport".
        ' Reason: if all the code is removed (deleted) in a component, this file does not export the changes.
        ' Then, in the next import, the code come back to component because the old file continues at 'src' folder.
        ' A fix to it could be delete the files at 'src' folder before export, but it is not recommended.
        ' Modified by Adriano Bortoloto https://github.com/AdrianoBortoloto Sep 16 2015
        'If hasCodeToExport(component) Then
            'Debug.Print "exporting type is " & component.Type
            Select Case component.Type
                Case vbext_ct_ClassModule
                    exportComponent export_path, component
                Case vbext_ct_StdModule
                    exportComponent export_path, component, ".bas"
                Case vbext_ct_MSForm
                    exportComponent export_path, component, ".frm"
                Case vbext_ct_Document
                    exportLines export_path, component
                Case Else
                    'Raise "Unkown component type"
            End Select
        'End If
        '''
    Next component
End Sub

Private Function hasCodeToExport(component As VBComponent) As Boolean
    hasCodeToExport = True
    If component.CodeModule.CountOfLines <= 2 Then
        Dim firstLine As String
        firstLine = Trim(component.CodeModule.lines(1, 1))
        'Debug.Print firstLine
        hasCodeToExport = Not (firstLine = "" Or firstLine = "Option Explicit")
    End If
End Function

'To export everything else but sheets
Private Sub exportComponent(exportPath As String, component As VBComponent, Optional extension As String = ".cls")
    Debug.Print "exporting " & component.name & extension

    '''
    '把 Version Control files 擋住不要匯出。
    Dim componentName As String
    componentName = component.name
    If IsItemInArray(componentName) Then
        Debug.Print "skip " & component.name & extension
        Exit Sub
    End If
    '''

    component.Export exportPath & "\" & component.name & extension
End Sub

'To export sheets
Private Sub exportLines(exportPath As String, component As VBComponent)
    Dim extension As String: extension = ".sheet.cls"
    Dim fileName As String
    fileName = exportPath & "\" & component.name & extension
    Debug.Print "exporting " & component.name & extension
    'component.Export exportPath & "\" & component.name & extension
    Dim FSO As New Scripting.FileSystemObject
    Dim outStream As TextStream
    Set outStream = FSO.CreateTextFile(fileName, True, False)
    
    '''
    ' If file do not have code, do not write in
    ' But in exportVbaCode() the componente must be exported even if it has no code.
    ' Thus, all future imports will pull the changes of components which code was full deleted
    ' avoiding pull old codes deleted before. See the Sub exportVbaCode()
    ' Modified by Adriano Bortoloto https://github.com/AdrianoBortoloto Sep 16 2015
    'outStream.Write (component.CodeModule.lines(1, component.CodeModule.CountOfLines))
    If Not component.CodeModule.CountOfLines = 0 Then
        outStream.Write (component.CodeModule.lines(1, component.CodeModule.CountOfLines))
    End If
    '''
    
    outStream.Close
End Sub


' Usually called after the given workbook is opened. The option includeClassFiles is False by default because
' they don't import correctly from VBA. They'll have to be imported manually instead.
Public Sub importVbaCode(VBAProject As VBProject, Optional includeClassFiles As Boolean = False)
    Dim vbProjectFileName As String
    On Error Resume Next
    'this can throw if the workbook has never been saved.
    vbProjectFileName = VBAProject.fileName
    On Error GoTo 0
    If vbProjectFileName = "" Then
        'In this case it is a new workbook, we skip it
        Debug.Print "No file name for project " & VBAProject.name & ", skipping"
        Exit Sub
    End If

    Dim export_path As String
    export_path = getSourceDir(vbProjectFileName, createIfNotExists:=False)
    If export_path = "" Then
        'The source directory does not exist, code has never been exported for this vbaProject.
        Debug.Print "No import directory for project " & VBAProject.name & ", skipping"
        Exit Sub
    End If

    'initialize globals for Application.OnTime
    Set componentsToImport = New Dictionary
    Set sheetsToImport = New Dictionary
    Set vbaProjectToImport = VBAProject

    Dim FSO As New Scripting.FileSystemObject
    Dim projContents As Folder
    Set projContents = FSO.GetFolder(export_path)
    Dim file As Object
    For Each file In projContents.Files()
        'check if and how to import the file
        checkHowToImport file, includeClassFiles
    Next

    Dim componentName As String
    Dim vComponentName As Variant
    'Remove all the modules and class modules
    For Each vComponentName In componentsToImport.keys
        componentName = vComponentName
        removeComponent VBAProject, componentName
    Next
    'Then import them
    Debug.Print "Invoking 'Build.importComponents'with Application.Ontime with delay " & IMPORT_DELAY
    ' to prevent duplicate modules, like MyClass1 etc.
    Application.OnTime Now() + TimeValue(IMPORT_DELAY), "'Build.importComponents'"
    Debug.Print "almost finished importing code for " & VBAProject.name
End Sub

Private Sub checkHowToImport(file As Object, includeClassFiles As Boolean)
    Dim fileName As String
    fileName = file.name
    '''
    '把 Version Control files 擋住不要匯入。
    Dim componentName As String
    componentName = Left(fileName, InStr(fileName, ".") - 1)
    If IsItemInArray(componentName) Then
        Debug.Print "skip " & fileName
        Exit Sub
    End If
    '''

    If Len(fileName) > 4 Then
        Dim lastPart As String
        lastPart = Right(fileName, 4)
        Select Case lastPart
            Case ".cls" ' 10 == Len(".sheet.cls")
                If Len(fileName) > 10 And Right(fileName, 10) = ".sheet.cls" Then
                    'import lines into sheet: importLines vbaProjectToImport, file
                    sheetsToImport.Add componentName, file
                Else
                    ' .cls files don't import correctly because of a bug in excel, therefore we can exclude them.
                    ' In that case they'll have to be imported manually.
                    If includeClassFiles Then
                        'importComponent vbaProject, file
                        componentsToImport.Add componentName, file.Path
                    End If
                End If
            Case ".bas", ".frm"
                'importComponent vbaProject, file
                componentsToImport.Add componentName, file.Path
            Case Else
                'do nothing
                Debug.Print "Skipping file " & fileName
        End Select
    End If
End Sub

' Only removes the vba component if it exists
Private Sub removeComponent(VBAProject As VBProject, componentName As String)
    If componentExists(VBAProject, componentName) Then
        Dim c As VBComponent
        Set c = VBAProject.VBComponents(componentName)
        Debug.Print "removing " & c.name
        VBAProject.VBComponents.Remove c
    End If
End Sub

Public Sub importComponents()
    If componentsToImport Is Nothing Then
        Debug.Print "Failed to import! Dictionary 'componentsToImport' was not initialized."
        Exit Sub
    End If
    Dim componentName As String
    Dim vComponentName As Variant
    For Each vComponentName In componentsToImport.keys
        componentName = vComponentName
        importComponent vbaProjectToImport, componentsToImport(componentName)
    Next

    'Import the sheets
    For Each vComponentName In sheetsToImport.keys
        componentName = vComponentName
        importLines vbaProjectToImport, sheetsToImport(componentName)
    Next

    Debug.Print "Finished importing code for " & vbaProjectToImport.name
    'We're done, clear globals explicitly to free memory.
    Set componentsToImport = Nothing
    Set vbaProjectToImport = Nothing
End Sub

' Assumes any component with same name has already been removed.
Private Sub importComponent(VBAProject As VBProject, filePath As String)
    Debug.Print "Importing component from  " & filePath
    'This next line is a bug! It imports all classes as modules!
    VBAProject.VBComponents.Import filePath
End Sub

Private Sub importLines(VBAProject As VBProject, file As Object)
    Dim componentName As String
    componentName = Left(file.name, InStr(file.name, ".") - 1)
    Dim c As VBComponent
    If Not componentExists(VBAProject, componentName) Then
        ' Create a sheet to import this code into. We cannot set the ws.codeName property which is read-only,
        ' instead we set its vbComponent.name which leads to the same result.
        Dim addedSheetCodeName As String
        addedSheetCodeName = addSheetToWorkbook(componentName, VBAProject.fileName)
        Set c = VBAProject.VBComponents(addedSheetCodeName)
        c.name = componentName
    End If
    Set c = VBAProject.VBComponents(componentName)
    Debug.Print "Importing lines from " & componentName & " into component " & c.name

    ' At this point compilation errors may cause a crash, so we ignore those.
    On Error Resume Next
    c.CodeModule.DeleteLines 1, c.CodeModule.CountOfLines
    c.CodeModule.AddFromFile file.Path
    On Error GoTo 0
End Sub

Public Function componentExists(ByRef proj As VBProject, name As String) As Boolean
    On Error GoTo doesnt
    Dim c As VBComponent
    Set c = proj.VBComponents(name)
    componentExists = True
    Exit Function
doesnt:
    componentExists = False
End Function

' Returns a reference to the workbook. Opens it if it is not already opened.
' Raises error if the file cannot be found.
Public Function openWorkbook(ByVal filePath As String) As Workbook
    Dim wb As Workbook
    Dim fileName As String
    fileName = Dir(filePath)
    On Error Resume Next
    Set wb = Workbooks(fileName)
    On Error GoTo 0
    If wb Is Nothing Then
        Set wb = Workbooks.Open(filePath) 'can raise error
    End If
    Set openWorkbook = wb
End Function

' Returns the CodeName of the added sheet or an empty String if the workbook could not be opened.
Public Function addSheetToWorkbook(sheetName As String, workbookFilePath As String) As String
    Dim wb As Workbook
    On Error Resume Next 'can throw if given path does not exist
    Set wb = openWorkbook(workbookFilePath)
    On Error GoTo 0
    If Not wb Is Nothing Then
        Dim ws As Worksheet
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.name = sheetName
        'ws.CodeName = sheetName: cannot assign to read only property
        Debug.Print "Sheet added " & sheetName
        addSheetToWorkbook = ws.CodeName
    Else
        Debug.Print "Skipping file " & sheetName & ". Could not open workbook " & workbookFilePath
        addSheetToWorkbook = ""
    End If
End Function

Public Function IsItemInArray(StrItem As String) As Boolean
    Const VerCtrlFilesSize = 10
    Dim VerCtrlFiles(VerCtrlFilesSize - 1) As String
    VerCtrlFiles(0) = "Build"
    VerCtrlFiles(1) = "ErrorHandling"
    VerCtrlFiles(2) = "Formatter"
    VerCtrlFiles(3) = "NamedRanges"
    VerCtrlFiles(4) = "Menu"
    VerCtrlFiles(5) = "Test"
    VerCtrlFiles(6) = "XMLexporter"
    VerCtrlFiles(7) = "CustomActions"
    VerCtrlFiles(8) = "EventListener"
    VerCtrlFiles(9) = "MyCustomActions"

    Dim i As Integer
    For i = LBound(VerCtrlFiles) To UBound(VerCtrlFiles)
        If StrComp(VerCtrlFiles(i), StrItem, vbTextCompare) = 0 Then
            IsItemInArray = True
            Exit Function
        End If
    Next i
    IsItemInArray = False
End Function

Sub TestIsItemInArray()
    Dim ItemToCheck As String
    ItemToCheck = "Build"
    
    If IsItemInArray(ItemToCheck) Then
        MsgBox ItemToCheck & " exists in the array."
    Else
        MsgBox ItemToCheck & " does not exist in the array."
    End If
End Sub
