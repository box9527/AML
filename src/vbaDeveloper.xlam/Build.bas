Attribute VB_Name = "Build"
'''
'這個檔案保留原本 https://github.com/hilkoc/vbaDeveloper.git#ac25532 的使用方式，並加入 PR#14 的部分修正。
'使用方式為將此目錄加入 Excel 的 Addins 的方式來使用，以便達到可以 Export/Import 又可以不會把這個目錄的檔案暴露出來的目的。
'但因為不符合我們的使用情境 (我們需要動態多人使用，無法使用絕對路徑的Addins)，故僅描述正確的使用方式並記錄於此。
'1.如果你是一個新的專案，請你打開一個新的 Excel，例如 A.xlsx
'2.因為步驟7 之後，很有可能會面臨檔案無法開啟的問題 (變成 xlam，Addin file)，所以請事先依照以下的目錄結構放置:
'  A.xlsx
'    |_src
'       |_vbaDeveloper.xlam
'          |_Build.bas (本檔案)
'          |_b.bas ...
'3.打開 A.xlsx， VB Scriptor (Alt+F11) 然後 檔案 (Menu) > 匯入檔案 (Import) (Ctrl+M)，將這個 Build.bas 檔案匯入。
'4.打開 VB Script Editor > Tools references (工具/設定引用項目) 檢查你是否將以下兩個方塊打勾:
'  * Microsoft Visual Basic for Applications Extensibility 5.x
'  * Microsoft Scripting Runtime
'5.將Project Name 改成 VBADeveloper (名字無法改變，除非你改程式不然會有Bug)，(Ctrl+R --> F4) 屬性--> Name。
'6.啟用對 VBA 的程式存取：
'  檔案 > 選項 > 信任中心 > 信任中心設定 > 巨集設定，
'  勾選此方塊：「啟用對 VBA 的程式存取」（在 excel 2010 中：「信任對 vba 項目物件模型的存取」）
'  如果您的政策設定不允許更改此選項，您可以建立以下註冊表​​項：
'  [HKEY_CURRENT_USER\Software\Policies\Microsoft\office\{Excel-Version}\excel\security]
'  “accessvbom”=dword:00000001
'  如果您在 Excel 2013 中收到「找不到路徑」異常，請執行下列步驟：
'  在“信任中心”設定中，前往“檔案封鎖設定”並取消選取“開啟”和/或“儲存”」對於「Excel 2007 及更高版本啟用巨集的工作簿和範本」。
'7.在專案 VBADeveloper/ThisWorkbook 檔案的屬性中 (F4) 找到 "IsAddin"，修改為 True。
'8.將檔案另存成 xlam 格式，如 vbaDeveloper.xlam，並關閉此檔。此時會變成如下的目錄結構:
'  vbaDeveloper.xlam
'    |_src
'      |_vbaDeveloper.xlam
'         |_Build.bas (本檔案)
'         |_b.bas ...  
'9.打開以你的專案的名字命名的 Excel，例如 poc_tool8.xlsx，並放在跟 vbaDeveloper.xlam同層位置:
'  poc_tool8.xlsx
'  vbaDeveloper.xlam
'    |_src
'      |_vbaDeveloper.xlam
'         |_Build.bas (本檔案)
'         |_b.bas ...
'10.在檔案 > 選項 > 增益集，在"執行"的地方按下後，利用瀏覽，匯入同層的 vbaDeveloper.xlam 作為增益集，確定後回到poc_tool8.xlsx。
'11.打開 VB Script Editor > Tools references (工具/設定引用項目) 將以下方塊打勾:
'   * VBADeveloper
'   沒有上述的方塊的話，就用 Tools references (工具/設定引用項目) > 瀏覽選取同一層的 vbaDeveloper.xlam 匯入。
'12.將Project Name 改成 NLPDeveloper，(Ctrl+R --> F4) 屬性--> Name。
'13.此時你應該會在左方看到 VBADeveloper 這個專案出現，選取 VBADeveloper 專案下方的模組 Build，你會看到本檔案。
'14.把你的游標移到 Build 的 testImport() 上，
'   然後 執行 > 執行 Sub 或 UserForm (F5)，就可以把 src/vbaDeveloper.xlsm 下相關的 VBA code 匯入。
'   如果以下三個檔案沒有自動匯入，則手動匯入:
'   * CustomActions.cls
'   * EventListener.cls
'   * MyCustomActions.cls
'15.在 NLPDeveloper 專案下的 ThisWookbook.cls ，新增以下兩個 function，用來使用 增益集 > NLPDeveloper > Import / Export:
'   Private Sub Workbook_Open()
'       menu.createMenu
'   End Sub
'   Private Sub Workbook_BeforeClose(Cancel As Boolean)
'       menu.deleteMenu
'   End Sub
'16.儲存你的 xlsx 成為 xlsm 格式，至此你的 poc_tool8.xlsm 應該就可以正常的使用匯進匯入VBA 功能。
'   需注意的是，這樣完成的 poc_tool8.xlsm 只能在你的電腦使用 (同層的 vbaDeveloper.xlam 不可以移動)，
'   不同電腦會出現找不到 Addin 的問題。這可能有辦法用程式解但沒有詳盡測試，請參考以下位置:
'   https://stackoverflow.com/questions/21350305/programmatically-install-add-in-vba
'''

Option Explicit

Private Const IMPORT_DELAY As String = "00:00:03"

'We need to make these variables public such that they can be given as arguments to application.ontime()
Public componentsToImport As Dictionary 'Key = componentName, Value = componentFilePath
Public sheetsToImport As Dictionary 'Key = componentName, Value = File object
Public vbaProjectToImport As VBProject

Public Sub testImport()
    Dim proj_name As String
    proj_name = "VBADeveloper"

    Dim VBAProject As Object
    Set VBAProject = Application.VBE.VBProjects(proj_name)
    Build.importVbaCode VBAProject
End Sub

Public Sub testExport()
    Dim proj_name As String
    proj_name = "VBADeveloper"

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
        'PR#14
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
    
    'PR#14
    'outStream.Write (component.codeModule.lines(1, component.codeModule.CountOfLines))
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
    Dim componentName As String
    componentName = Left(fileName, InStr(fileName, ".") - 1)
    If componentName = "Build" Then
        '"don't remove or import ourself
        Exit Sub
    End If

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
