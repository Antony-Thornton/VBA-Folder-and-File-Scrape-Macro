# VBA-Folder-and-File-Scrape-Macro
An excel based VBA Macro to pull details of folders and files with the option to delete specific files.


## Business Case
At Alphabet, each department was tasked with reviewing any files that may contain data that needs to be anonymised or deleted due to GDPR/DPA regulations. This required all of the teams to pick through all files and folders that may be subject to this. 

## Solution
To help speed up the process and flag files that may be subject to quick deletion I created a VBA macro that scraped through all of the folders and files and extracted the information into an excel data frame. The user could then initially assess the file based on the information provided and select those that can be deleted by putting yes in a field. Another macro would then use the information collected and the yes/no field to loop through all the folders and delete them. 

## Outcome
The macro was initially created within my team to help us but I expanded it to make it company friendly. In one case a department no longer had to hire a temp to review folders and files. Others saved countless man hours and I won an award for the macro. 

### Sample code
     

    Public Sub NonRecursiveMethod()

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

        Dim fso, oFolder, oSubfolder, oFile, queue As Collection
        Dim FldrPicker As FileDialog

        Set fso = CreateObject("Scripting.FileSystemObject")
        Set queue = New Collection



        Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

        With FldrPicker
           .AllowMultiSelect = False
           .Show
           MyFolder = .SelectedItems(1)
           Err.Clear
        End With

        Set Folder = fso.GetFolder(MyFolder)

        queue.Add fso.GetFolder(MyFolder)

        Do While queue.Count > 0
            Set oFolder = queue(1)
            queue.Remove 1 'dequeue
            '...insert any folder processing code here...
            For Each oSubfolder In oFolder.subfolders
                queue.Add oSubfolder 'enqueue
            Next oSubfolder
            For Each oFile In oFolder.Files
                '...insert any file processing code here...
                '==================================================================================================

                '==================================================================================================
            Next oFile
        Loop
        Dim myUserForm As UserForm1
    Set myUserForm = New UserForm1
    myUserForm.Show


    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationManual

    End Sub


