import sys
import os
import tkinter as tk
from tkinter import messagebox

def ok_button_click():
    selected_group = listbox.get(listbox.curselection())
    group = ""
    if selected_group == "Hydra_ING":
        group = ".\\settings\\Hydra_ING.lst"
    elif selected_group == "Hydra_ING_Active_directory":
        group = ".\\settings\\Hydra_ING_Active_directory.lst"
    elif selected_group == "Hydra_ING_Common_Tasks":
        group = ".\\settings\\Hydra_ING_Common_Tasks.lst"
    elif selected_group == "Hydra_ING_SCCM":
        group = ".\\settings\\Hydra_ING_SCCM.lst"
    elif selected_group == "Hydra_ING_Virtualization":
        group = ".\\settings\\Hydra_ING_Virtualization.lst"
    else:
        messagebox.showerror("Error", "Invalid group selected")
        return
    
    sequences_list_param = group
    sequences_list = [file.name for file in os.scandir(".\\settings") if file.name.endswith(".lst")]
    sequences_list = list(set(sequences_list))
    
    # Optional script parameter defining an alternative Sequence List
    sequences_list_param = input("Enter the alternative Sequence List: ")
    
    def add_object_list_to_grid(object_list, file_path):
        for obj in object_list:
            obj = obj.strip()
            if len(obj) == 0:
                continue
            row_id = output_data_grid.insert("", "end", values=(obj, "0", "Pending", "Pending", 0, "", file_path, True, 0, "-"))
            output_data_grid.set(row_id, 7, False)
            output_data_grid.item(row_id, tags=("White",))
            object_options = {
                "GroupID": "0",
                "PreviousStateComment": "",
                "StepProtocol": None,
                "SharedVariable": None
            }
            output_data_grid.set(row_id, 0, tags=object_options)
    
    def cancel_all_force():
        ids_to_cancel = []
        for row_index in output_data_grid.selection():
            ids_to_cancel.append(output_data_grid.item(row_index)["values"][8])
        ids_to_cancel = list(set(ids_to_cancel))
        
        for row_index in output_data_grid.get_children():
            if output_data_grid.item(row_index)["values"][4] == 0 or output_data_grid.item(row_index)["values"][4] == -5:
                state_to_set = -6
                output_data_grid.set(row_index, 0, "#")
                output_data_grid.set(row_index, 1, "Cancelled")
                output_data_grid.set(row_index, 2, "CANCELLED")
                output_data_grid.set(row_index, 3, state_to_set)
                output_data_grid.set(row_index, 7, True)
                output_data_grid.item(row_index, tags=("CANCELLED",))
                if output_data_grid.item(row_index)["values"][0].tags["GroupID"] == "0":
                    output_data_grid.set(row_index, 0, True)
                output_data_grid.item(row_index)["values"][0].tags["StepProtocol"] = "\r\n" + output_data_grid.item(row_index)["values"][0] + "   Cancelled - Not started"
            if output_data_grid.item(row_index)["values"][4] > 0:
                state_to_set = -6
                output_data_grid.set(row_index, 0, "#")
                output_data_grid.set(row_index, 1, "Cancelled")
                output_data_grid.set(row_index, 2, "CANCELLED")
                output_data_grid.set(row_index, 3, state_to_set)
                output_data_grid.set(row_index, 7, True)
                output_data_grid.item(row_index, tags=("CANCELLED",))
                if output_data_grid.item(row_index)["values"][0].tags["GroupID"] == "0":
                    output_data_grid.set(row_index, 0, True)
                output_data_grid.item(row_index)["values"][0].tags["StepProtocol"] = "\r\n" + output_data_grid.item(row_index)["values"][0] + "   Cancelled"
    
    # Create the form
    form = tk.Tk()
    form.title("Select ADgroup you want to add to this user")
    form.geometry("300x200")
    form.resizable(False, False)
    
    # Create the OK button
    ok_button = tk.Button(form, text="OK", width=10, command=ok_button_click)
    ok_button.place(x=75, y=120)
    
    # Create the Cancel button
    cancel_button = tk.Button(form, text="Cancel", width=10, command=form.destroy)
    cancel_button.place(x=150, y=120)
    
    # Create the label
    label = tk.Label(form, text="Please select a group:")
    label.place(x=10, y=20)
    
    # Create the listbox
    listbox = tk.Listbox(form, height=5)
    listbox.place(x=10, y=40)
    listbox.insert(tk.END, "Hydra_ING")
    listbox.insert(tk.END, "Hydra_ING_Active_directory")
    listbox.insert(tk.END, "Hydra_ING_Common_Tasks")
    listbox.insert(tk.END, "Hydra_ING_SCCM")
    listbox.insert(tk.END, "Hydra_ING_Virtualization")
    
    form.mainloop()











root.mainloop()



# Reset the schedulers, if any
for ID in IDsToCancel:
    if Sequences[ID].SequenceScheduler != 0:
        # An object with a scheduler has been cancelled: all relative objects will be cancelled too
        Sequences[ID].SequenceScheduler = 0
        Sequences[ID].SequenceSchedulerExpired = True
        for item in OutputDataGridSequence[ID]:
            StateToSet = -6
            SetCellValue(GridIndex, item.Index, "#", "Cancelled", "CANCELLED", StateToSet, Colors.Get_Item("CANCELLED"), "#", "#")
            item.Cells[7].Value = FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked
            item.Cells[7].ReadOnly = False
            item.Cells[0].Tag.StepProtocol = "\r\n" + OutputDataGrid.Rows[RowIndex].Cells[0].Value
            item.Cells[0].Tag.StepProtocol += "   Cancelled - Not started"
            item.DefaultCellStyle.BackColor = Colors.Get_Item("CANCELLED")

def Cancel_Sequence(RowsSelected):
    # Cancel the Sequences of the objects passed as argument
    IDsToCancel = list(set(OutputDataGrid.Rows[RowIndex].Cells[8].Value for RowIndex in RowsSelected))
    for RowIndex in RowsSelected:
        # Set the Cells Values to "Cancel", re-enable the checkbox
        if OutputDataGrid.Rows[RowIndex].Cells[4].Value == 0:
            # Sequence not started
            StateToSet = -6
            SetCellValue(GridIndex, RowIndex, "#", "Cancelled", "CANCELLED", StateToSet, Colors.Get_Item("CANCELLED"), "#", "#")
            OutputDataGrid.Rows[RowIndex].Cells[7].Value = FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked
            OutputDataGrid.Rows[RowIndex].Cells[7].ReadOnly = False
            OutputDataGrid.Rows[RowIndex].Cells[0].Tag.StepProtocol = "\r\n" + OutputDataGrid.Rows[RowIndex].Cells[0].Value
            OutputDataGrid.Rows[RowIndex].Cells[0].Tag.StepProtocol += "   Cancelled - Not started"
            OutputDataGrid.Rows[RowIndex].DefaultCellStyle.BackColor = Colors.Get_Item("CANCELLED")
        if OutputDataGrid.Rows[RowIndex].Cells[4].Value > 0:
            # Sequence started
            StateToSet = -5
            SetCellValue(GridIndex, RowIndex, "#", "Cancelling", "CANCELLING", StateToSet, Colors.Get_Item("CANCELLED"), "#", "#")
            OutputDataGrid.Rows[RowIndex].DefaultCellStyle.BackColor = Colors.Get_Item("CANCELLED")
    # Reset the schedulers, if any
    for ID in IDsToCancel:
        if Sequences[ID].SequenceScheduler != 0:
            # An object with a scheduler has been cancelled: all relative objects will be cancelled too
            Sequences[ID].SequenceScheduler = 0
            Sequences[ID].SequenceSchedulerExpired = True
            for item in OutputDataGridSequence[ID]:
                StateToSet = -6
                SetCellValue(GridIndex, item.Index, "#", "Cancelled", "CANCELLED", StateToSet, Colors.Get_Item("CANCELLED"), "#", "#")
                item.Cells[7].Value = FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked
                item.Cells[7].ReadOnly = False
                item.Cells[0].Tag.StepProtocol = "\r\n" + OutputDataGrid.Rows[RowIndex].Cells[0].Value
                item.Cells[0].Tag.StepProtocol += "   Cancelled - Not started"
                item.DefaultCellStyle.BackColor = Colors.Get_Item("CANCELLED")

def clear_grid():{}
OutputDataGrid.Rows.Clear()
OutputDataGrid.Rows[0].Cells[7].Value = False
OutputDataGrid.Rows[0].Cells[7].ReadOnly = True
nbCheckedBoxes = 0
ObjectsLabel.Text = ""
OutputDataGridContextMenuObject.Items[5].Text = ""
SequencesTreeView.SelectedNode = SequencesTreeView.Nodes[0]
get_count_checkboxes()

CreateRunspace = lambda RunspaceScriptBlock, RunspaceArg, RSId, SharedVariable: {}
    # Create a Runspace for the Object passed as RunspaceArg, with the code RunspaceScriptBlock
Powershell = PowerShell.Create().AddScript(RunspaceScriptBlock).AddArgument(RunspaceArg.strip()).AddArgument(SharedVariable)
Powershell.RunspacePool = RunspacePool[RSId]
RunspaceCollection[RSId].append({
        "Runspace": Powershell.BeginInvoke(),
        "PowerShell": Powershell
    })

def Export_Result(ExportFormat):
    # Close the Export form and start the right Export, based on the format given as parameter
    if ExportFormat == 0:
        Export_ToCSV()
    elif ExportFormat == 1:
        Export_ToExcel()
    elif ExportFormat == 2:
        Export_ToHTML(True, True)
    elif ExportFormat == 3:
        Send_Email()
        
        
        
        
        
        
        
        
        
def Export_ToCSV():{}
Selection = []
if FormExportColCheckBox[1].Checked:
    Selection.append(0)
if FormExportColCheckBox[2].Checked:
    Selection.append(2)
if FormExportColCheckBox[3].Checked:
    Selection.append(3)
    New_Item(CSVTempPath, "File", "Force")
    if FormExportHeaderCheckBox.Checked:
        ToPaste = " ".join([OutputDataGrid.Columns[i].Name for i in Selection])
        if FormExportColCheckBox[4].Checked:
            ToPaste = ToPaste + " " + "Sequence Name"
        Add_Content(CSVTempPath, ToPaste)
    if FormExportSelectionCheckBox.Checked == True:
        OutputGridRows = [OutputDataGrid.Rows[cell.RowIndex] for cell in OutputDataGrid.SelectedCells]
    else:
        OutputGridRows = OutputDataGrid.Rows[0:OutputDataGrid.RowCount - 1]
    for Row in OutputGridRows:
        ToPaste = " ".join([Row.Cells[i].EditedFormattedValue for i in Selection])
        if FormExportColCheckBox[4].Checked:
            ToPaste = ToPaste + " " + Sequences[Row.Cells[8].EditedFormattedValue].SequenceLabel
        Add_Content(CSVTempPath, ToPaste)
    Start_Process('C:\windows\system32\notepad.exe', CSVTempPath)

def Export_ToExcel():
    Export_ToHTML(False, False)
    Cc = threading.thread.CurrentThread.CurrentCulture
    threading.thread.CurrentThread.CurrentCulture = 'en-US'
    Excel = New_Object(ComObject, "Excel.Application")
    Excel.Visible = True
    WorkBook = Excel.Workbooks.Open(HTMLTempPath)
    Excel.Windows.Item(1).Displaygridlines = True
    Now = Get_Date().ToString("yyyyMMddHHssmm")
    NewXLSXName = (New_Object(System.IO.FileInfo, Split_Path(XLSXTempPath, "Leaf")).BaseName + "_" + Now + (New_Object(System.IO.FileInfo, Split_Path(XLSXTempPath, "Leaf")).Extension))
    NewXLSXPath = Join_Path(Split_Path(XLSXTempPath, "Parent"), NewXLSXName)
    Workbook.SaveAs(NewXLSXPath)
    threading.thread.CurrentThread.CurrentCulture = Cc

def Export_CreateHTML(Object, TaskResult, Step, SequenceName, Color, OnlySelection, WithStyle):
    GridObjects = []
    ColumnSelection = []
    if Color:
        ColumnSelection.append('Color')
    if Object:
        ColumnSelection.append(OutputDataGrid.Columns[0].Name)
    if TaskResult:
        ColumnSelection.append("Task*")
    if Step:
        ColumnSelection.append(OutputDataGrid.Columns[3].Name)
    if SequenceName:
        ColumnSelection.append('Sequence Name')
    if OnlySelection:
        OutputGridRows = [OutputDataGrid.Rows[cell.RowIndex] for cell in OutputDataGrid.SelectedCells]
    else:
        OutputGridRows = OutputDataGrid.Rows[0:OutputDataGrid.RowCount - 1]
    for Row in OutputGridRows:
        Prop = {}
        if Object:
            Prop[OutputDataGrid.Columns[0].Name] = Row.Cells[0].EditedFormattedValue
        if TaskResult and FormExportHeaderCheckBox.Checked:
            Prop[OutputDataGrid.Columns[2].Name] = Row.Cells[2].EditedFormattedValue
        if TaskResult and not FormExportHeaderCheckBox.Checked:
            i = 0
            for item in Row.Cells[2].EditedFormattedValue.split(";"):
                i += 1
                Prop["TaskResultPart " + str(i)] = item
        if Step:
            Prop[OutputDataGrid.Columns[3].Name] = Row.Cells[3].EditedFormattedValue
        if SequenceName:
            if Sequences[Row.Cells[8].EditedFormattedValue].SequenceLabel != None:
                Prop['Sequence Name'] = Sequences[Row.Cells[8].EditedFormattedValue].SequenceLabel
            else:
                Prop['Sequence Name'] = " "
        if Color:
            ColorHex = Row.Cells[5].EditedFormattedValue.replace("#FF", "#")
            Prop['Color'] = "###" + ColorHex + "###"
        obj = New_Object(PSObject, Property=Prop)
        GridObjects.append(obj)
    if WithStyle:
        HTMLStyle = "<style>"
        HTMLStyle += "body { background-color:#dddddd; font-family:Tahoma; font-size:12pt; }"
        HTMLStyle += "td, th { border:1px solid black; border-collapse:collapse; }"
        HTMLStyle += "th { color:white; background-color:black; }"
        HTMLStyle += "table, tr, td, th { padding: 2px; margin: 0px }"
        HTMLStyle += "table { margin-left:50px; }"
        HTMLStyle += "</style>"
        
        
        
        
        
        
    else:
        HTMLStyle = ""
HTMLBody = GridObjects.loc[:, ColumnSelection].to_html(index=False)  # Create the HTML code filtering the columns
if not FormExportHeaderCheckBox.Checked:
    # The Header shouldn't be displayed
    HTMLBody = re.sub(r"<tr><th>.*?</th></tr>", "", HTMLBody)  # Suppress the Automatic Header
if Color:
    # String manipulation to get the color at the correct place in the HTML code
    HTMLBody = re.sub(r"><td>###", " bgcolor=", HTMLBody)
    HTMLBody = re.sub(r"###</td>", ">", HTMLBody)
    HTMLBody = re.sub(r"<th>Color</th>", "", HTMLBody)

return HTMLBody, HTMLStyle

def Export_Group(GroupExported):
    # Export Group(s) to a XML bundle
    if GroupExported == "All Groups":
        # Set the GroupList file: it will contain links to the single XML Group files
        SaveFileDialog = tkinter.filedialog.asksaveasfile(initialdir=LastDirExportGroup, filetypes=[("Hydra Group List", "*.grouplist"), ("All files", "*.*")])
        LastDirExportGroup = os.path.dirname(SaveFileDialog.name)
        if SaveFileDialog.name == "":
            return
        with open(SaveFileDialog.name, "w") as file:
            file.write("# Hydra Group List\n")  # Create the file with a header
        GroupsInCurrentGrid = GridObjects.loc[GridObjects["GroupID"] != 0, "GroupID"].unique().tolist()  # Enumerate the groups in the grid
        for GroupUsedItem in GroupsInCurrentGrid:
            # Loop in the list of Groups to export
            if GroupUsedItem == 0:
                continue  # Skip the value 0 that doesn't belong to any Group
            ExportFileName = os.path.join(os.path.dirname(SaveFileDialog.name), f"{GroupUsedItem}.group.xml")  # The name automatically generated are using the name of the Groups themselves
            Export_GroupSingle(GroupUsedItem, ExportFileName)  # Call the Export_GroupSingle function to export GroupUsedItem and save it into ExportFileName
            with open(SaveFileDialog.name, "a") as file:
                file.write(f"{GroupUsedItem}.group.xml\n")  # Add the path of the newly created XML in the GroupList file
    else:
        # Export one Group only
        Export_GroupSingle(GroupExported, "")

def Export_GroupSingle(GroupToExport, FileForExport):
    # Create an XML file with all attributes needed for a re-import
    SeqId = 0
    ObjectsList = []
    for i in range(OutputDataGrid.RowCount - 1):
        if OutputDataGrid.Rows[i].Cells[0].Tag.GroupID == GroupToExport:
            # Extract the member of the group GroupToExport
            SeqID = OutputDataGrid.Rows[i].Cells[8].Value  # Get the corresponding Sequence ID
            ObjectsList.append(OutputDataGrid.Rows[i].Cells[0].Value)  # Create a list with the Group members
    if SeqID == 0:
        return  # SeqID=0, No group found
    if FileForExport == "":
        # Get the name for saving if not given as parameter
        SaveFileDialog = tkinter.filedialog.asksaveasfile(initialdir=LastDirExportGroup, filetypes=[("Hydra Group", "*.group.xml"), ("All files", "*.*")])
        LastDirExportGroup = os.path.dirname(SaveFileDialog.name)  # Save the last folder for registry settings saving on close
        if SaveFileDialog.name == "":
            return
        FileForExport = SaveFileDialog.name
    # Save the different attributes arrays in Export variables
    ScriptBlockExport = Sequences[SeqID].ScriptBlock
    ScriptBlockCommentExport = Sequences[SeqID].ScriptBlockComment
    ScriptBlockVariableExport = Sequences[SeqID].ScriptBlockVariable
    ScriptBlockModuleExport = Sequences[SeqID].ScriptBlockModule
    ScriptBlockCheckboxesExport = Sequences[SeqID].ScriptBlockCheckboxes
    ScriptBlockPreLoadExport = Sequences[SeqID].ScriptBlockPreLoad
    SequenceSchedulerExport = Sequences[SeqID].SequenceScheduler
    MaxThreadsExport = Sequences[SeqID].MaxThreads
    MaxCheckedObjectsExport = Sequences[SeqID].MaxCheckedObjects
    SequenceSendMailExport = Sequences[SeqID].SequenceSendMail
    SequenceLabelExport = Sequences[SeqID].SequenceLabel
    BelongsToGroupExport = Sequences[SeqID].BelongsToGroup
    SequenceSchedulerExpiredExport = Sequences[SeqID].SequenceSchedulerExpired
    SecurityCodeExport = Sequences[SeqID].SecurityCode
    DisplayWarningExport = Sequences[SeqID].DisplayWarning
    ExportVer = 1  # File Version number
    GroupExported = GroupToExport  # Export the variables set to FileForExport with the command Export_Clixml
    with open(FileForExport, "wb") as file:
        pickle.dump([ExportVer, ObjectsList, GroupExported, ScriptBlockExport, ScriptBlockCommentExport, ScriptBlockPreLoadExport, ScriptBlockVariableExport, ScriptBlockModuleExport, ScriptBlockCheckboxesExport, MaxThreadsExport, SequenceSchedulerExport, MaxCheckedObjectsExport, SequenceSendMailExport, SequenceLabelExport, BelongsToGroupExport, SequenceSchedulerExpiredExport, SecurityCodeExport, DisplayWarningExport], file)
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
def Export_Tabs():
    SaveFileDialog = System.Windows.Forms.SaveFileDialog()
    SaveFileDialog.InitialDirectory = LastDirExportGroup
    SaveFileDialog.Filter = "Hydra Tabs (*.tabs)|*.tabs|All files|*.*"
    SaveFileDialog.ShowDialog().OutNull()
    LastDirExportGroup = os.path.split(SaveFileDialog.FileName)
    if SaveFileDialog.FileName == "":
        return
    with open(SaveFileDialog.FileName, "w") as file:
        file.write("# Hydra Tabs\n")
        for i in range(DataGridTabControl.TabCount):
            TabGridID = DataGridTabControl.TabPages[i].Tag.TabPageIndex
            TabText = DataGridTabControl.TabPages[i].Text
            TabColor1 = DataGridTabControl.TabPages[i].Tag.ColorSelected
            TabColor2 = DataGridTabControl.TabPages[i].Tag.ColorUnSelected
            TabSave = f"{TabText};{TabColor1};{TabColor2};{';'.join([cell.Value for cell in OutputDataGridTab[TabGridID].Rows.Cells if cell.ColumnIndex == 0])}"
            file.write(TabSave + "\n")

def Export_ToHTML(WithStyle, OpenFile):
    HTMLExport = Export_CreateHTML(FormExportColCheckBox[1].Checked, FormExportColCheckBox[2].Checked, FormExportColCheckBox[3].Checked, FormExportColCheckBox[4].Checked, FormExportColorCheckBox.Checked, FormExportSelectionCheckBox.Checked, WithStyle)
    if WithStyle:
        with open(HTMLTempPath, "w") as file:
            file.write(f"<H2>Sequence Results</H2> {HTMLExport[0]}")
    else:
        with open(HTMLTempPath, "w") as file:
            file.write(HTMLExport[0])
    if OpenFile:
        os.startfile(HTMLTempPath)

def Start_NewRunspaceScriptBlock(Row, RSId):
    ConcurrentJobs[RSId] += 1
    if Row.Cells[4].Value == 0:
        Row.Cells[0].Tag.StepProtocol.append(f"\n{Row.Cells[0].Value} - {Sequences[RSId].SequenceLabel}\n   Started at {CurrentTime}")
    if Sequences[RSId].ScriptBlockCheckboxes[Row.Cells[4].Value].Checked:
        CreateRunspace(Sequences[RSId].ScriptBlock[Row.Cells[4].Value], Row.Cells[0].Value, RSId, Row.Cells[0].Tag.SharedVariable)
    else:
        CreateRunspace("return \"OK\", \"{} (Step {} Skipped)\"".format(Row.Cells[0].Tag.PreviousStateComment, Row.Cells[4].Value+1), Row.Cells[0].Value, RSId, Row.Cells[0].Tag.SharedVariable)
    Row.Cells[4].Value += 1
    JobNb[RSId] += 1
    SetCellValue(SequenceTabIndex[RSId], Row.index, JobNb[RSId], "Executing task: {}".format(Sequences[RSId].ScriptBlockComment[Row.Cells[4].Value - 1]), "Step {}".format(Row.Cells[4].Value), Row.Cells[4].Value, "#", "#", "#")
    
    
    
    
    
    
    
    
    
    
GetData = lambda: {}
CurrentTime = datetime.datetime.now().strftime("%H:%M:%S")
StillRunning = False
TotalConcurrentJobs = 0
GroupsRunning = []   
for Item in SequencesToParse:
    if Sequences[Item].SequenceScheduler != 0:
        TimeDiff = datetime.datetime.now() - Sequences[Item].SequenceScheduler
        TimeDiffFormatted = "{:02d}:{:02d}:{:02d}".format(TimeDiff.seconds // 3600, (TimeDiff.seconds // 60) % 60, TimeDiff.seconds % 60)
        if TimeDiff.total_seconds() <= 1:
            Sequences[Item].SequenceScheduler = 0
            Sequences[Item].SequenceSchedulerExpired = True
        
        for Row in OutputDataGridSequence[Item]:
            if not Row.Cells[7].Value:
                continue
            if Sequences[Item].SequenceScheduler != 0:
                Row.Cells[3].Value = TimeDiffFormatted
            
            if Row.Cells[4].Value > 0 or Row.Cells[4].Value == -5:
                GetRowState(Row, SequenceTabIndex[Item])
                if Row.Cells[4].Value > 0 or Row.Cells[4].Value == -5:
                    StillRunning = True
                    if Row.Cells[0].Tag.GroupID not in GroupsRunning:
                        GroupsRunning.append(Row.Cells[0].Tag.GroupID)
    
    for Item in SequencesToParse:
        if Sequences[Item].SequenceScheduler != 0:
            StillRunning = True
            continue
        if ConcurrentJobs[Item] > RunspacePool[Item].GetMaxRunspaces():
            continue
        for Row in OutputDataGridSequence[Item]:
            if Item != Row.Cells[8].Value or not Row.Cells[7].Value:
                continue
            if Row.Cells[4].Value >= 0 and "Executing" not in Row.Cells[2].Value:
                StillRunning = True
                if Row.Cells[0].Tag.GroupID not in GroupsRunning:
                    GroupsRunning.append(Row.Cells[0].Tag.GroupID)
                if Row.Index == 0:
                    if "PreLoad" in Sequences[Item].ScriptBlockCheckboxes[Row.Cells[4].Value].Text:
                        Sequences[Item].ScriptBlockPreLoad = True
                        StartNewRunspaceScriptBlock(Row, Item)
                        Sequences[Item].ScriptBlockCheckboxes[Row.Cells[4].Value - 1].Checked = False
                    elif "PreLoad" not in Sequences[Item].ScriptBlockCheckboxes[Row.Cells[4].Value].Text:
                        Sequences[Item].ScriptBlockPreLoad = False
                if Sequences[Item].ScriptBlockPreLoad:
                    break
                StartNewRunspaceScriptBlock(Row, Item)
                if ConcurrentJobs[Item] > RunspacePool[Item].GetMaxRunspaces():
                    break
    
    for Item in SequencesToParse:
        TotalConcurrentJobs += ConcurrentJobs[Item] - 1
    
    ObjectsLabel.Text = "Total Objects: {}, Running: {}, Done: {}".format(OutputDataGrid.RowCount - 1, TotalConcurrentJobs, ObjectsDone)
    RunningTask = len([row for row in OutputDataGrid.Rows if int(row.Cells[4].Value) > 0])
    MenuToolStrip.Items.FirstOrDefault(lambda x: x.ToolTipText == "Clear the Grid").Enabled = RunningTask == 0
    
    if not StillRunning:
        SetSequenceFinished()

def GetCountCheckboxes():
    nbCheckedBoxes = len([row for row in OutputDataGrid.Rows if row.Cells[7].Value])
    SetObjectsState()
    SetActionButtonState()

def GetFileButton(NameFilter, InitialPath):
    OpenFileDialog = System.Windows.Forms.OpenFileDialog()
    if InitialPath != "":
        OpenFileDialog.InitialDirectory = InitialPath
    OpenFileDialog.Filter = NameFilter
    OpenFileDialog.ShowHelp = True
    OpenFileDialog.ShowDialog()
    return OpenFileDialog.FileName

def GetIPRange():
        # Create a list of IP's
    import ipaddress

IPRangeFromText = "192.168.0.1"
IPRangeToText = "192.168.0.10"

try:
    ipaddress.ip_address(IPRangeFromText)
    ipaddress.ip_address(IPRangeToText)
except ValueError:
    # One of the values entered is not a correct IP
    messagebox.showerror("Error", "Unable to validate the IP.")
    return

#IP operations
IP1 = ipaddress.ip_address(IPRangeFromText).packed[::-1]
IP1 = ipaddress.ip_address(".".join(str(byte) for byte in IP1)).packed
IP2 = ipaddress.ip_address(IPRangeToText).packed[::-1]
IP2 = ipaddress.ip_address(".".join(str(byte) for byte in IP2)).packed

# Create the IP range
IPObjectList = []
for x in range(int.from_bytes(IP1, 'big'), int.from_bytes(IP2, 'big')+1):
    IP = ipaddress.ip_address(x.to_bytes(4, 'big')).packed[::-1]
    IPObjectList.append(".".join(str(byte) for byte in IP))

if len(IPObjectList) == 0:
    messagebox.showerror("Error", "Unable to create a range with these values.")
    return

if len(IPObjectList) > 16384:
    messagebox.showerror("Error", "Unable to create a range with more than 16384 values.")
    return

FormIPRange.Close()
Add_ObjectListToGrid(IPObjectList, "")  # Add the objects created to the grid

def Get_NewSequenceList():
    # Load a new Sequence List
    SequencesListPath = Get_FileButton("Hydra Sequence List (*.lst)|*.lst|All files|*.*", HydraSettingsPath)
    if SequencesListPath == None or SequencesListPath == "":
        return
    Set_ReloadSequenceList()

def Get_ObjectsAD():
    # Get AD objects from a Query
    import ActiveDirectory

    if not ActiveDirectory:
        # Load the AD module if not already loaded
        import ActiveDirectory

    # Run the Query defined in the "Query AD" window
    ADObjList = eval(ADQueryText.Text)
    ADObjList = [obj for obj in ADObjList if obj.startswith(ADQueryFilterText.Text)]

    if ADObjList == None:
        messagebox.showinfo("AD Query", "Nothing found matching your query")
        return

    FormADQuery.Close()
    Add_ObjectListToGrid(ADObjList, "")  # Add the objects found to the grid

def Get_ObjectsFile():
    # Read the content of a file and add every object to the grid
    ObjFilePath = Get_FileButton("All files (*.*)|*.*", LastDirObjects)
    if ObjFilePath == None or ObjFilePath == "":
        return

    Separator = [",", ";", "|"]
    ObjectList = [obj.strip() for obj in open(ObjFilePath).read().split(Separator) if obj.strip() != ""]

    Add_ObjectListToGrid(ObjectList, ObjFilePath)  # Add the objects found to the grid as well as their corresponding file
    LastDirObjects = os.path.split(ObjFilePath)[0]  # Save the last directory for registry user's settings on close

def Get_ObjectsManually():
    # Enter or paste a list of objects separated by separators
    ObjectList = Read_InputBoxDialog("Objects", "Enter the list of Objects separated by a comma, semicolon or pipe:", "")
    if ObjectList == "":
        return

    Separator = [",", ";", "|"]
    ObjectList = [obj.strip() for obj in ObjectList.split(Separator) if obj.strip() != ""]

    Add_ObjectListToGrid(ObjectList, "")  # Add the objects to the grid

def Get_ObjectsPatse():
    # Paste the objects stored in the Clipboard to the grid
    Clipboard = tkinter.clipboard_get()
    if Clipboard == "":
        return

    Separator = [",", ";", "|", "\r", "\n", "\t"]
    Clipboard = [obj.strip() for obj in Clipboard.split(Separator) if obj.strip() != ""]

    ObjectList = [item for item in Clipboard if item != ""]

    Add_ObjectListToGrid(ObjectList, "")  # Add the objects to the grid

def Get_ObjectsSCCM(QueryTpye):
    # Get Objects with a SCCM Query
    QueryTpye = QueryTpye.lower()

    if QueryTpye == "object":
        ObjectPattern = SCCMQueryObjText.Text
        WmiQuery = f"""
            Select DISTINCT *
            FROM SMS_R_System
            WHERE SMS_R_System.Name IS LIKE '{ObjectPattern}%'
        """
    elif QueryTpye == "ip":
        IPPattern = SCCMQueryIPText.Text
        WmiQuery = f"""
            Select DISTINCT *
            FROM SMS_R_System
            WHERE SMS_R_System.IPAddresses IS LIKE '{IPPattern}%'
        """
    elif QueryTpye == "manual":
        WmiQuery = SCCMQueryManualText.Text

    WmiParams = {
        'ComputerName': SCCM_ConfigMgrSiteServer,
        'Namespace': f"root\sms\site_{SCCM_SiteCode}",
        'Query': WmiQuery
    }

    try:
        # Execute the Query
        SCCMQueryResult = Get_WmiObject(**WmiParams).select("Name")
    except:
        messagebox.showinfo("SCCM Query", "Unable to connect to the SCCM server")
        return
    if SCCMQueryResult == "":
        MessageBox.Show("Nothing found matching your query", "SCCM Query", MessageBoxButtons.OK, MessageBoxIcon.Stop)
    return
FormSCCMQuery.Close()
Add_ObjectListToGrid(SCCMQueryResult, "")
def Get_RegistrySettings():
    RegHydra = Get_ItemProperty ("HKCU:\SOFTWARE\Hydra\5, ErrorAction=SilentlyContinue").select
    if RegHydra != None:
        for RegEntry in RegHydra:
            if RegEntry.Name.startswith("Color_"):
                Colors[RegEntry.Name.split("_")[1]] = RegEntry.Value
            else:
                globals()[RegEntry.Name] = RegEntry.Value
def Get_RowState(row, TabIndex):
    xPID = row.Cells[1].Value
    SeqId = row.Cells[8].Value
    try:
        if RunspaceCollection[SeqId][xPID].Runspace.IsCompleted:
            Get_RowState_ReturnedValue(row, xPID)
            if row.Cells[4].Value == -5:
                row.Cells[4].Value = -6
                row.Cells[2].Value = "Cancelled"
                row.Cells[0].Tag.PreviousStateComment = "Cancelled"
                row.Cells[0].Tag.StepProtocol.append("   Cancelled")
                row.Cells[3].Value = "CANCELLED"
                row.DefaultCellStyle.BackColor = row.Cells[5].Value
                if row.Cells[0].Tag.GroupID == "0":
                    row.Cells[0].ReadOnly = False
                row.Cells[7].Value = FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked
                row.Cells[7].ReadOnly = False
                if row.Cells[7].Value == False:
                    nbCheckedBoxes -= 1
                ObjectsDone += 1
            if row.Cells[4].Value == len(Sequences[SeqId].ScriptBlockCheckboxes):
                row.Cells[4].Value = -1
                row.DefaultCellStyle.BackColor = row.Cells[5].Value
                if row.Cells[0].Tag.GroupID == "0":
                    row.Cells[0].ReadOnly = False
                row.Cells[7].Value = FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked
                row.Cells[7].ReadOnly = False
                if row.Cells[7].Value == False:
                    nbCheckedBoxes -= 1
                ObjectsDone += 1
            if row.Cells[4].Value < 0:
                row.Cells[0].Tag.StepProtocol.append("   Ended at " + CurrentTime)
            if Sequences[SeqId].ScriptBlockPreLoad:
                for i in range(1, len(OutputDataGridSequence[SeqId])):
                    OutputDataGridSequence[SeqId][i].Cells[0].Tag.SharedVariable = row.Cells[0].Tag.SharedVariable
                    if row.Cells[4].Value < -1:
                        OutputDataGridSequence[SeqId][i].Cells[4].Value = -6
                        OutputDataGridSequence[SeqId][i].Cells[2].Value = "Cancelled - PreLoad not OK"
                        OutputDataGridSequence[SeqId][i].Cells[0].Tag.PreviousStateComment = "Cancelled"
                        OutputDataGridSequence[SeqId][i].Cells[0].Tag.StepProtocol.append("   Cancelled")
                        OutputDataGridSequence[SeqId][i].Cells[3].Value = "CANCELLED"
                        OutputDataGridSequence[SeqId][i].DefaultCellStyle.BackColor = Colors["CANCELLED"]
                        if OutputDataGridSequence[SeqId][i].Cells[0].Tag.GroupID == "0":
                            OutputDataGridSequence[SeqId][i].Cells[0].ReadOnly = False
                        OutputDataGridSequence[SeqId][i].Cells[7].Value = FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked
                        OutputDataGridSequence[SeqId][i].Cells[7].ReadOnly = False
                        if OutputDataGridSequence[SeqId][i].Cells[7].Value == False:
                            nbCheckedBoxes -= 1
                        ObjectsDone += 1
    except:
        row.Cells[4].Value = -6
        row.Cells[2].Value = "Runspace Error - Cancelled"
        row.Cells[0].Tag.PreviousStateComment = "Runspace Error - Cancelled"
        row.Cells[0].Tag.StepProtocol.append("   Runspace Error - Cancelled")
        row.Cells[3].Value = "CANCELLED"
        row.Cells[5].Value = Colors["STOP"]
        row.DefaultCellStyle.BackColor = row.Cells[5].Value
        row.Cells[0].Tag.StepProtocol.append("   Ended at " + CurrentTime)
        if row.Cells[0].Tag.GroupID == "0":
            row.Cells[0].ReadOnly = False
        row.Cells[7].Value = FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked
        row.Cells[7].ReadOnly = False
        if row.Cells[7].Value == False:
            nbCheckedBoxes -= 1
        ObjectsDone += 1
        
        
        
        
        
        
        
        
        
        
        
        
        def Get_RowState_ReturnedValue(row, xPID):
            SeqId = row.Cells[8].Value
try:
        xReceive = RunspaceCollection[SeqId][xPID].PowerShell.EndInvoke(RunspaceCollection[SeqId][xPID].Runspace)
        RunspaceCollection[SeqId][xPID].PowerShell.Dispose()
        JobResultState = xReceive[0]
        JobResultComment = xReceive[1]
except:
        JobResultState = "ERROR"
        JobResultComment = "Error in Task (Not enough objects returned: enable Debug for details)"
        if DebugMode == 5:
            print("\n", row.Cells[0].Value, ": Not enough Objects returned")
            Write_DebugReceiveOutput(xReceive)
        xReceive = ["ERROR", "", Colors.Get_Item("CANCELLED")]
        row.Cells[4].Value = -4
if len(xReceive) > 4:
        JobResultState = "ERROR"
        JobResultComment = "Error in Task (Too much objects returned: enable Debug for details)"
        row.Cells[4].Value = -4
        if DebugMode == 5:
            print("\n", row.Cells[0].Value, ": ", len(xReceive), " Objects returned (too much)")
            Write_DebugReceiveOutput(xReceive)
        xReceive = ["ERROR", "", Colors.Get_Item("CANCELLED")]
elif xReceive[0] not in ["OK", "BREAK", "STOP", "ERROR"]:
        JobResultState = "ERROR"
        JobResultComment = "Error in Task (Wrong keyword returned: enable Debug for details)"
        row.Cells[4].Value = -4
        if DebugMode == 5:
            print("\n", row.Cells[0].Value, ": Wrong keyword returned: ", xReceive[0])
            Write_DebugReceiveOutput(xReceive)
        xReceive = ["ERROR", "", Colors.Get_Item("CANCELLED")]
if xReceive[0] == "STOP" and row.Cells[4].Value >= 0:
        JobResultState = "STOP at step ", row.Cells[4].Value
        row.Cells[4].Value = -2
if xReceive[0] == "BREAK" and row.Cells[4].Value >= 0:
        JobResultState = "BREAK at step ", row.Cells[4].Value
        row.Cells[4].Value = -3
if len(xReceive) >= 3 and xReceive[2] != None:
        ColorsReturned = xReceive[2].split("|")
        try:
            windows.media.color(ColorsReturned[0])
            IsHTMLColor = True
        except:
            IsHTMLColor = False
        if re.match('#ff(([0-9a-f]{6}))\b', ColorsReturned[0]) or IsHTMLColor:
            BackgroundColor = ColorsReturned[0]
        else:
            BackgroundColor = Colors.Get_Item(xReceive[0])
            if DebugMode == 5 and ColorsReturned[0] != "":
                print("\nWrong color value: ", ColorsReturned[0], " is not a valid HTML color or in format #FFxxxxxx. Reset to default")
        if len(ColorsReturned) > 1:
            FontStyles = ["Regular", "Italic", "Bold", "Strikeout", "Underline"]
            if ColorsReturned[1] in FontStyles:
                row.Cells[2].Style.Font = Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle(ColorsReturned[1]))
            else:
                CellFontReturned = ColorsReturned[1].split(',')
                try:
                    row.Cells[2].Style.Font = Drawing.Font(CellFontReturned)
                except:
                    row.Cells[2].Style.Font = Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Regular)
                    if DebugMode == 5:
                        print("\nWrong Font value: ", ColorsReturned[1], " is not a valid System.Drawing.Font. Reset to default")
                        
                        
                        
                        
                        
                        
                        
if len(ColorsReturned) > 2:
    try:
        ColorsReturned[2] = ColorsReturned[2].lstrip('#')
        int(ColorsReturned[2], 16)
        IsHTMLColor = True
    except:
        IsHTMLColor = False
    if re.match(r'#ff(([0-9a-f]{6}))\b', ColorsReturned[2]) or IsHTMLColor:
        row.Cells[2].Style.ForeColor = ColorsReturned[2]
else:
    BackgroundColor = Colors.Get_Item(xReceive[0])
if len(xReceive) == 4:
    row.Cells[0].Tag.SharedVariable = xReceive[3]
if row.Cells[4].Value > -5:
    row.Cells[2].Value = JobResultComment
    if "Skipped" not in JobResultComment:
        row.Cells[0].Tag.PreviousStateComment = JobResultComment
    row.Cells[0].Tag.StepProtocol.append("   " + JobResultComment)
    row.Cells[3].Value = JobResultState
    row.Cells[5].Value = BackgroundColor
if row.Cells[4].Value < 0:
    row.DefaultCellStyle.BackColor = row.Cells[5].Value
    row.Cells[7].Value = FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked
    row.Cells[7].ReadOnly = False
    if row.Cells[7].Value == False:
        nbCheckedBoxes -= 1
    if row.Cells[0].Tag.GroupID == "0":
        row.Cells[0].ReadOnly = False
    nbCheckedBoxes -= 1
    ObjectsDone += 1
ConcurrentJobs[SeqId] -= 1
if ConcurrentJobs[SeqId] == 0:
    ConcurrentJobs[SeqId] = 1

def Get_Sequence(FileSeqPath, SeqName):
    SequencePanelCheckbox = []
    SequencePanelLabel = []
    SequencePanelVariable = []
    SequenceTasksPanel.Controls.Clear()
    TimerInterval = TimerIntervalDefault
    Timer.Interval = TimerInterval
    if not os.path.exists(FileSeqPath):
        Set_SequencePanelTitle(SeqName + "\n\n\n", "Red")
        Set_SequencePanelLabel("  Missing: " + FileSeqPath + " not found\n\n", "Italic", "Red", 0)
        SequenceName = ""
        ActionButton.Enabled = False
        return
    try:
        xmldata = ET.parse(FileSeqPath)
    except ET.ParseError:
        Set_SequencePanelTitle(SeqName + "\n\n\n", "Red")
        Set_SequencePanelLabel("  Error: XML parse error in " + FileSeqPath + "\n\n", "Italic", "Red", 0)
        ActionButton.Enabled = False
        return
    SequencePath = ""
    SequenceAbsolutePath = ""
    SecurityCode = ""
    FileSeqParentPath = os.path.dirname(os.path.abspath(FileSeqPath))
    Err = False
    ScriptBlockLoaded = []
    ScriptBlockCommentLoaded = []
    MaxThreadsText.Text = DefaultThreads
    MaxObj = 0
    DisplayWarning = False
    SendMail = False
    MailServer = ""
    MailFrom = ""
    MailTo = ""
    MailReplyTo = ""
    SeqPosition = 0
    XMLSeqParam = xmldata.findall(".//parameter")
    for XMLParam in XMLSeqParam:
        if XMLParam.attrib["name"] == "sequencename":
            SeqName = XMLParam.attrib["value"]
        elif XMLParam.attrib["name"] == "warning":
            if XMLParam.attrib["value"] == "yes":
                DisplayWarning = True
        elif XMLParam.attrib["name"] == "securitycode":
            SecurityCode = XMLParam.attrib["value"]
        elif XMLParam.attrib["name"] == "maxthreads":
            MaxThreadsText.Text = XMLParam.attrib["value"]
        elif XMLParam.attrib["name"] == "maxobjects":
            MaxObj = XMLParam.attrib["value"]
        elif XMLParam.attrib["name"] == "sendmail":
            if XMLParam.attrib["value"] == "yes":
                SendMail = True
        elif XMLParam.attrib["name"] == "timer":
            if int(XMLParam.attrib["value"]) >= 500 and int(XMLParam.attrib["value"]) <= 30000:
                TimerInterval = int(XMLParam.attrib["value"])
                Timer.Interval = TimerInterval    
if XMLParam.name.startswith("mail"):
    globals()[XMLParam.name] = XMLParam.value
XMLSeqMod = xmldata.sequence.importmodule
SequenceImportModuleLoaded = []
for XMLModule in XMLSeqMod:
    if XMLModule.type == "ImportPSSnapIn" or XMLModule.type == "ImportPSModulesFromPath" or XMLModule.type == "ImportPSModule":
        SequenceImportModuleLoaded.append({
            "Type": XMLModule.type,
            "Name": XMLModule.name
        })        
XMLSeqVar = xmldata.sequence.variable
SeqVariablesPos = []
for XMLVar in XMLSeqVar:
    if XMLVar.type in VariableTypes:
        SeqVariablesPos.append({
            "Type": XMLVar.type,
            "Name": XMLVar.name,
            "Value": XMLVar.value
        })

XMLSeqPreload = xmldata.sequence.preload
for XMLPreload in XMLSeqPreload:
    SeqPath = XMLPreload.path
    SeqComment = XMLPreload.comment
    SeqFound = False
    SeqLocation = ""
    TaskRelativeTo = os.path.join(FileSeqParentPath, SeqPath)
    if os.path.exists(TaskRelativeTo):
        SeqFound = True
        SeqLocation = TaskRelativeTo
    elif os.path.exists(SeqPath):
        SeqFound = True
        SeqLocation = SeqPath
    if SeqFound:
        ScriptBlockLoaded.append(get_command(SeqLocation).ScriptBlock)
        if len(Error) != 0:
            ErrorMsg = Error[0].split("\n")[0].split(".ps1:")[1]
            Set_SequencePanelCheckbox("PreLoad " + str(SeqPosition+1) + "\n", SeqPosition)
            Set_SequencePanelLabel("  Error: error detected Line:" + ErrorMsg + "\n", "Italic", "Red", SeqPosition)
            Err = True
        else:
            ScriptBlockCommentLoaded.append(SeqComment)
            Set_SequencePanelCheckbox("PreLoad " + str(SeqPosition+1) + "\n", SeqPosition)
            Set_SequencePanelLabel("  " + SeqComment + "\n\n", "Regular", "Magenta", SeqPosition)
    else:
        Set_SequencePanelCheckbox("PreLoad " + str(SeqPosition+1) + "\n", SeqPosition)
        Set_SequencePanelLabel("  Missing: " + SeqPath + "\n\n", "Italic", "Red", SeqPosition)
        Err = True
    SeqPosition += 1

XMLSeqTask = xmldata.sequence.task
for XMLTask in XMLSeqTask:
    SeqPath = XMLTask.path
    SeqComment = XMLTask.comment
    SeqFound = False
    SeqLocation = ""
    TaskRelativeTo = os.path.join(FileSeqParentPath, SeqPath)
    if os.path.exists(TaskRelativeTo):
        SeqFound = True
        SeqLocation = TaskRelativeTo
    elif os.path.exists(SeqPath):
        SeqFound = True
        SeqLocation = SeqPath
    if SeqFound:
        ScriptBlockLoaded.append(get_command(SeqLocation).ScriptBlock)
        if len(Error) != 0:
            ErrorMsg = Error[0].split("\n")[0].split(".ps1:")[1]
            Set_SequencePanelCheckbox("Step " + str(SeqPosition+1) + "\n", SeqPosition)
            Set_SequencePanelLabel("  Error: error detected Line:" + ErrorMsg + "\n", "Italic", "Red", SeqPosition)
            Err = True
        else:
            ScriptBlockCommentLoaded.append(SeqComment)
            Set_SequencePanelCheckbox("Step " + str(SeqPosition+1) + "\n", SeqPosition)
            Set_SequencePanelLabel("  " + SeqComment + "\n\n", "Regular", "Green", SeqPosition)
    else:
        Set_SequencePanelCheckbox("Step " + str(SeqPosition+1) + "\n", SeqPosition)
        Set_SequencePanelLabel("  Missing: " + SeqPath + "\n\n", "Italic", "Red", SeqPosition)
        Err = True
    SeqPosition += 1
    
    
    
    
    
    
    
    
SequenceName = SeqName
SequencePath = os.path.dirname(FileSeqPath)
SequenceAbsolutePath = FileSeqParentPath
SequenceFullPath = FileSeqPath
MaxSteps = len(ScriptBlockLoaded)  # Set the number of Steps of the loaded sequence
MaxObjects = MaxObj
Set-ActionButtonState()

def Get_SequenceFileManual():
    SeqFilePath = Get_FileButton("Hydra Sequence (*.sequence.xml)|*.sequence.xml|All files|*.*", LastDirSequences)
    
    if SeqFilePath is None or SeqFilePath == "":
        return
    
    LastDirSequences = os.path.dirname(SeqFilePath)  # Set the variable LastDirSequences to the folder of the sequence choosen. This will be reused as default folder for the next manual load
    ManuallyLoadedSeq = next((node for node in SequencesTreeView.Nodes if node.Name == "*Manually Loaded*"), None)  # Check if the Sequences Tree already has a parent node "Manually Loaded"
    
    if ManuallyLoadedSeq is None:
        # The parent node "Manually Loaded" doesn't exist and is created
        SequenceListRootNode = System.Windows.Forms.TreeNode()
        SequenceListRootNode.Text = "Manually Loaded"
        SequenceListRootNode.Name = "Manually Loaded"
        SequencesTreeView.Nodes.Add(SequenceListRootNode)  # Add the new parent node "Manually Loaded" at the bottom of the Sequence tree
        properties = {'SeqName': "----- Manually Loaded -----", 'SeqPath': ""}
        object = PSObject(properties)
        SequenceList.append(object)  # The properties of this new parent node are added in the SequenceList array
    
    SequenceListRootNode = next((node for node in SequencesTreeView.Nodes if node.Name == "*Manually Loaded*"), None)  # Connects to "Manually Loaded"
    SequenceListSubNode = System.Windows.Forms.TreeNode()
    xmldata = System.Xml.XmlDocument()  # Creates a new XML object
    SeqName = SeqFilePath
    
    try:
        # Search for a paramater "sequencename" in the .sequence.xml file selected
        xmldata.Load(os.path.abspath(SeqFilePath))
        for parameter in xmldata.sequence.parameter:
            if parameter.Name == "sequencename":
                SeqName = f" {parameter.Value} ({SeqFilePath})"
    except System.Xml.XmlException:
        # if it's not found, the name of the Sequence will be the path of the file
        pass
    
    SequenceListSubNode.Text = SeqName
    SequenceListSubNode.Tag = len(SequencesTreeView.Nodes.Nodes) + len(SequencesTreeView.Nodes)                   
    SequenceListRootNode.Nodes.Add(SequenceListSubNode)  # Add the new node SequenceListSubNode in the "Manually Loaded" section
    properties = {'SeqName': SeqName, 'SeqPath': SeqFilePath}
    object = PSObject(properties)
    SequenceList.append(object)  # The properties of this new node are added in the SequenceList array
    
    SequencesTreeView.SelectedNode = SequenceListSubNode  # Select this new node
    
    if FormSettingsSequenceExpandedRadioButton.Checked:
        # Depending on the user's settings, expand or collapse the Sequence Tree  
        SequenceListRootNode.Expand()
    else:
        SequenceListRootNode.Collapse()
    
    SelectionChanged = True

def Import_Group():
    FileToImport = Get_FileButton("Hydra Groups Files(*.group.xml,*.grouplist)|*.group.xml;*.grouplist|Hydra Group (*.group.xml)|*.group.xml|Hydra Group List (*.grouplist)|*.grouplist|All files (*.*)|*.*", LastDirImportGroup)
    
    if FileToImport == "" or FileToImport is None:
        return
    
    LastDirImportGroup = os.path.dirname(FileToImport)  # Set the variable LastDirImportGroup to the folder of the sequence choosen. This will be reused as default folder for the next import
    ErrorReturn = []
    NonImported = []
    
    if "Hydra Group List" in open(FileToImport).readline():
        # The file loaded is a grouplist
        for Path in open(FileToImport):
            # Parse the file and import all files contained in it
            Path = os.path.join(LastDirImportGroup, Path)
            
            if "Hydra Group List" in Path:
                continue
            
            if os.path.exists(Path):
                ErrorReturn.append(Import_GroupSingle(Path, True))  # If an import fails, the ErrorReturn won't be None
            else:
                NonImported.append(Path)                           
        else:
    # A single group file has been chosen
         Import_GroupSingle(FileToImport, False)
    SequencesTreeView.SelectedNode = SequencesTreeView.Nodes[0]  # Reset the sequence
    SelectionChanged = True
    SequenceLoaded = False
    return
if ErrorReturn is not None:
    # ErrorReturn is not empty: some groups to import were already existing
     MessageBox.Show("The following Groups are already assigned and have been skipped:\r\n\r\n" + "\r\n".join(ErrorReturn), "Groups", MessageBoxButtons.OK, MessageBoxIcon.Stop)

if NonImported is not None:
    # NonImported is not empty: some groups to import were not found
    MessageBox.Show("The following Groups could not be found:\r\n\r\n" + "\r\n".join(NonImported), "Groups", MessageBoxButtons.OK, MessageBoxIcon.Stop)

SequencesTreeView.SelectedNode = SequencesTreeView.Nodes[0]  # Reset the sequence
SelectionChanged = True
SequenceLoaded = False

def Import_GroupSingle(FileToImport, Silent):
    # Loads an XML file and set all attributes to the corresponding group
    for item in Import_Clixml(FileToImport):
        globals()[item.Name] = item.Value

    if GroupExported in GroupsUsed:
        # Check if the Group name found in the file is already in use
        if not Silent:
            MessageBox.Show("The Group Name '" + GroupExported + "' is already assigned.", "Group Name", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        return GroupExported  # Returns the name of the group to add in ErrorReturn in the function Import-Group

    # The Group name found does not exist: increase all arrays needed for a new Sequence and add the content of the respective variables
    NewSeq = {}
    NewSeq["ScriptBlock"] = ScriptBlockExport
    NewSeq["ScriptBlockComment"] = ScriptBlockCommentExport
    NewSeq["ScriptBlockPreLoad"] = False
    NewSeq["ScriptBlockVariable"] = ScriptBlockVariableExport
    NewSeq["ScriptBlockModule"] = ScriptBlockModuleExport
    NewSeq["ScriptBlockCheckboxes"] = ScriptBlockCheckboxesExport
    NewSeq["SequenceScheduler"] = SequenceSchedulerExport
    NewSeq["SequenceSendMail"] = SequenceSendMailExport
    NewSeq["MaxCheckedObjects"] = MaxCheckedObjectsExport
    NewSeq["SequenceLabel"] = SequenceLabelExport
    NewSeq["MaxThreads"] = MaxThreadsExport
    NewSeq["BelongsToGroup"] = BelongsToGroupExport
    NewSeq["SequenceSchedulerExpired"] = SequenceSchedulerExpiredExport
    NewSeq["SecurityCode"] = SecurityCodeExport
    NewSeq["DisplayWarning"] = DisplayWarningExport
    Sequences.append(NewSeq)
    RunspaceCollection.append([])
    RunspacePool.append([])
    JobNb.append([])
    ConcurrentJobs.append([])
    SequenceTabIndex += OutputDataGrid.Tag.TabPageIndex
    OutputDataGridSequence.append([])
    GroupsUsed.append(GroupExported)

DataGridViewCellStyleBold = System.Windows.Forms.DataGridViewCellStyle()
DataGridViewCellStyleBold.Alignment = 16
DataGridViewCellStyleBold.Font = System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Bold, 3, 0)
DataGridViewCellStyleBold.ForeColor = System.Drawing.Color.FromArgb(255, 0, 0, 0)
DataGridViewCellStyleBold.SelectionBackColor = System.Drawing.Color.FromArgb(255, 51, 153, 255)
DataGridViewCellStyleBold.SelectionForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)

DataGridViewCellStyleRegular = System.Windows.Forms.DataGridViewCellStyle()
DataGridViewCellStyleRegular.Alignment = 16
DataGridViewCellStyleRegular.Font = System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Regular, 3, 0)
DataGridViewCellStyleRegular.ForeColor = System.Drawing.Color.FromArgb(255, 0, 0, 0)
DataGridViewCellStyleRegular.SelectionBackColor = System.Drawing.Color.FromArgb(255, 51, 153, 255)
DataGridViewCellStyleRegular.SelectionForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)

if isinstance(Sequences[-1]["SequenceScheduler"], datetime):
    # A scheduler has been set
    TimeDiff = datetime.now() - Sequences[-1]["SequenceScheduler"]
    if TimeDiff.total_seconds() <= 1:
        # Timer expired
        Sequences[-1]["SequenceScheduler"] = 0
        PendingText = "Pending"
    else:
        PendingText = Sequences[-1]["SequenceScheduler"].strftime("%H:%M:%S")
        
        
        
        
        
        
        
        
        
        
 # No timer defined
PendingText = "Pending"
SeqId = len(Sequences) - 1
for Obj in ObjectsList:
    RowID = OutputDataGrid.Rows.Add(Obj, "0", f"Sequence assigned: {SequenceLabelExport}", PendingText, 0, "", "", True, SeqId, f"{GroupExported} ({MaxThreadsExport})")
    OutputDataGrid.Rows[RowID].Cells[0].Style = DataGridViewCellStyleBold
    OutputDataGrid.Rows[RowID].Cells[2].Style = DataGridViewCellStyleRegular
    OutputDataGrid.Rows[RowID].Cells[0].ReadOnly = True  # Objects names can't be modified in Groups
    ObjectOptions = PSObject()
    ObjectOptions.GroupID = GroupExported
    ObjectOptions.PreviousStateComment = ""
    ObjectOptions.StepProtocol = None
    ObjectOptions.SharedVariable = None
    OutputDataGrid.Rows[RowID].Cells[0].Tag = ObjectOptions
    OutputDataGrid.Rows[RowID].Cells[7].ReadOnly = False
    OutputDataGrid.Rows[RowID].DefaultCellStyle.BackColor = "White"
    OutputDataGridSequence[SeqId] += OutputDataGrid.Rows[RowID]  # Add the row defined in the OutputDataGridSequence of the current Sequence
SequenceLoaded = True
GetCountCheckboxes()
return None

# Import Tab(s) from a .tabs file
FileToImport = GetFileButton("Hydra Tabs (*.tabs)|*.tabs|All files (*.*)|*.*", LastDirImportGroup)
if FileToImport == "" or FileToImport == None:
    return None
LastDirImportGroup = os.path.split(FileToImport)[0]  # Set the variable LastDirImportGroup to the folder of the sequence choosen. This will be reused as default folder for the next import
for Line in open(FileToImport):
    if "Hydra Tabs" in Line:
        continue  # Skip the header
    LineSplit = Line.split(";")  # Read and set the different Tab attributes
    TabName = LineSplit[0]
    TabColor1 = LineSplit[1]
    TabColor2 = LineSplit[2]
    Objects = LineSplit[3:]
    SetNewTab()
    LastTab = DataGridTabControl.TabCount - 1
    GridTabIndex = DataGridTabControl.TabPages[LastTab].Tag.TabPageIndex
    DataGridTabControl.TabPages[LastTab].Text = f"  {TabName}  "
    DataGridTabControl.TabPages[LastTab].Tag.ColorSelected = TabColor1
    DataGridTabControl.TabPages[LastTab].Tag.ColorUnSelected = TabColor2
    OutputDataGrid = OutputDataGridTab[GridTabIndex]
    AddObjectListToGrid(Objects, "")
DataGridTabControl.SelectedIndex = DataGridTabControl.TabCount - 1
OutputDataGrid.ClearSelection()

# Remove a Tab
RunningTask = [row for row in OutputDataGrid.Rows if int(row.Cells[4].Value) > 0]  # Check if some objects are in a runing state and exits if any
if len(RunningTask) > 0:
    MessageBox.Show("Unable to remove a Tab while Sequences are running.\n", "Tab", MessageBoxButtons.OK, MessageBoxIcon.Stop)
    return None
TabToDelete = DataGridTabControl.SelectedTab
ClearGrid()
DataGridTabControl.TabPages.Remove(TabToDelete)

# Rename the Tab
NewTabName = ReadInputBoxDialog("Tab", "Set the new Tab Name:")
if NewTabName == "":
    return None
DataGridTabControl.TabPages[TabIndex].Text = f"  {NewTabName}  "

# Recreate the Objects list of the selected tab with the Objects name only
AllObjects = []
for i in range(OutputDataGrid.RowCount - 1):
    AllObjects.append(OutputDataGrid.Rows[i].Cells[0].Value)  # Get the Objects Name only
AllObjectsFile = []
for i in range(OutputDataGrid.RowCount - 1):
    AllObjectsFile.append(OutputDataGrid.Rows[i].Cells[6].Value)  # Get the Objects' corresponding files they belong
ClearGrid()  # Clear the grid
AddObjectListToGrid(AllObjects, "")  # Recreate the list with the prior saved names
for i in range(OutputDataGrid.RowCount - 2):
    OutputDataGrid.Rows[i].Cells[6].Value = AllObjectsFile[i]  # Add the corresponding files names

# Reset all settings to Default
ReallyReset = MessageBox.Show("Do you really want to reset all settings to the default values ?\nThis will close this session of Hydra.", "WARNING", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
if ReallyReset == "yes":
    os.remove("HKCU:\Software\Hydra\5")  # Delete and recreate HKCU:\Software\Hydra\5
    os.makedirs("HKCU:\Software\Hydra\5")
    SetDefaultSettings()  # Set the Settings to default
    ResetSettings = True  # With this option, Hydra won't save anything in the Registry on exit
    Form.Close()
    def Reset_Runspaces():
    # Parse the sequence previously ran and close and dispose their RunspacePool
     for Index in SequencesToParse:
        try:
            RunspacePool[Index].Close()
            RunspacePool[Index].Dispose()
            RunspaceCollection[Index] = []
            RunspacePool[Index] = []
        except:
            pass
    # Clear some memory parts using the Garbage Collection
    import gc
    gc.collect()
    # Set the SequencesToParse to an empty array to prepare the next sequences run
    SequencesToParse = []

def Reset_SequenceArrays():
    # Empty Sequence arrays that aren't used anymore
    # Detect all Sequence ID's on all Tabs
    SequenceIndex = []
    for i in range(DataGridTabControl.TabCount):
        TabGridID = DataGridTabControl.TabPages[i].Tag.TabPageIndex
        SequenceIndex += [cell.Value for cell in OutputDataGridTab[TabGridID].Rows.Cells if cell.ColumnIndex == 8]
    
    if SequenceIndex == None:
        return  # No Sequence found
    
    AllSequencesId = list(range(1, len(Sequences)))
    SequenceArraysToDelete = list(set(AllSequencesId) - set(SequenceIndex))  # Match the difference between the Sequences ID's assigned and the ID's in the grids
    for item in SequenceArraysToDelete:
        # The difference is the Sequences not assigned to any object anymore: all corresponding arrays can be emptied
        try:
            Sequences[item] = None
            OutputDataGridSequence[item] = None
        except:
            pass

def Restore_CurrentSettings():
    # Reset the user's settings as they were if the user cancels the changes he made
    for i in range(5):
        FormSettingsPathsText[i].Text = CurrentSettings[i]
    FormSettingsLogCheckBox.Checked = CurrentSettings[5]
    for i in range(4):
        FormSettingsColorsButton[i].BackColor = CurrentSettings[i + 6]
    FormSettingsColorsGUIBackLabel.BackColor = CurrentSettings[10]
    FormSettingsColorsGUIBackButton.BackColor = CurrentSettings[11]
    FormSettingsSplashScreenCheckBox.Checked = CurrentSettings[12]
    FormSettingsDebugScreenCheckBox.Checked = CurrentSettings[13]
    FormSettingsSequenceExpandedRadioButton.Checked = CurrentSettings[14]
    FormSettingsColorsGUIBackButton.BackColor = CurrentSettings[15]
    FormSettingsColorsGUISeqButton.BackColor = CurrentSettings[16]
    FormSettingsColorsGUISeqRunButton.BackColor = CurrentSettings[17]
    FormSettingsSequenceShowSearchRadioButton.Checked = CurrentSettings[18]
    FormSettingsSequenceShowHideRadioButton.Checked = CurrentSettings[19]
    FormSettingsSplashScreenCheckBox.Checked = CurrentSettings[20]
    FormSettingsDebugScreenCheckBox.Checked = CurrentSettings[21]
    FormSettingsSequenceExpandedRadioButton.Checked = CurrentSettings[22]
    FormSettingsSequenceCollapsedRadioButton.Checked = CurrentSettings[23]
    FormSettingsRowHeaderGroupBoxHiddenRadioButton.Checked = CurrentSettings[24]
    FormSettingsRowHeaderGroupBoxVisibleRadioButton.Checked = CurrentSettings[25]
    FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked = CurrentSettings[26]
    FormSettingsCheckBoxesGroupBoxResetStateRadioButton.Checked = CurrentSettings[27]
    for i in range(4):
        FormSettingsMailText[i].Text = CurrentSettings[i + 28]
    FormSettingsGroupsWarningUncheckedRadioButton.Checked = CurrentSettings[32]
    FormSettingsGroupsWarningCheckedRadioButton.Checked = CurrentSettings[33]
    FormSettingsGroupsThreadsVisibleRadioButton.Checked = CurrentSettings[34]
    FormSettingsGroupsThreadsInvisibleRadioButton.Checked = CurrentSettings[35]

def Save_CurrentSettings():
    # Save the state of the user's settings in case the user will cancel the process
    CurrentSettings = [FormSettingsPathsText[0].Text, FormSettingsPathsText[1].Text, FormSettingsPathsText[2].Text, FormSettingsPathsText[3].Text,
                       FormSettingsPathsText[4].Text, FormSettingsLogCheckBox.Checked, FormSettingsColorsButton[0].BackColor, FormSettingsColorsButton[1].BackColor,
                       FormSettingsColorsButton[2].BackColor, FormSettingsColorsButton[3].BackColor, FormSettingsColorsGUIBackLabel.BackColor, FormSettingsColorsGUIBackButton.BackColor,
                       FormSettingsSplashScreenCheckBox.Checked, FormSettingsDebugScreenCheckBox.Checked, FormSettingsSequenceExpandedRadioButton.Checked,
                       FormSettingsColorsGUIBackButton.BackColor, FormSettingsColorsGUISeqButton.BackColor, FormSettingsColorsGUISeqRunButton.BackColor,
                       FormSettingsSequenceShowSearchRadioButton.Checked, FormSettingsSequenceShowHideRadioButton.Checked, FormSettingsSplashScreenCheckBox.Checked,
                       FormSettingsDebugScreenCheckBox.Checked, FormSettingsSequenceExpandedRadioButton.Checked, FormSettingsSequenceCollapsedRadioButton.Checked,
                       FormSettingsRowHeaderGroupBoxHiddenRadioButton.Checked, FormSettingsRowHeaderGroupBoxVisibleRadioButton.Checked, FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked,
                       FormSettingsCheckBoxesGroupBoxResetStateRadioButton.Checked, FormSettingsMailText[0].Text, FormSettingsMailText[1].Text, FormSettingsMailText[2].Text,
                       FormSettingsMailText[3].Text, FormSettingsGroupsWarningUncheckedRadioButton.Checked, FormSettingsGroupsWarningCheckedRadioButton.Checked,
                       FormSettingsGroupsThreadsVisibleRadioButton.Checked, FormSettingsGroupsThreadsInvisibleRadioButton.Checked]

def Send_Email():
    # Send the state of the grid per email
    if EMailSMTPServer == "" or EMailSendFrom == "" or EMailSendTo == "" or EMailReplyTo == "":
        # If parameters are missing, exit
        import tkinter.messagebox as messagebox
        messagebox.showerror("Error", "Unable to find the e-mail parameters.\n\nEnter the parameters in the Settings panel.")
        return
    
    
    
    
    
    
    
    
    
    
    # Create a HMTL content based on the options checked by the user, and add the header
HTMLExport = Export_CreateHTML(FormExportColCheckBox[1].Checked, FormExportColCheckBox[2].Checked, FormExportColCheckBox[3].Checked, FormExportColCheckBox[4].Checked, FormExportColorCheckBox.Checked, FormExportSelectionCheckBox.Checked, True)
ToSend = ConvertToHtml(Head=HTMLExport[1], Body="<H2>Hydra Deployment Results</H2> " + HTMLExport[0]).OutString()
fSendMail(EMailSMTPServer, EMailSendFrom, EMailSendTo, EMailReplyTo, "Hydra Deployment Results", ToSend, True)

# Parse the finished Sequences and send an email if it was set in the sequence.xml file
MailToSend = False
for Item in SequencesToParse:
    if Sequences[Item].SequenceSendMail:
        if not MailToSend:
            MailToSend = True
            OutputDataGrid.ClearSelection()
        for row in OutputDataGridSequence[Item]:
            row.Selected = True
if MailToSend:
    HTMLExport = Export_CreateHTML(True, True, True, True, True, True, True)
    ToSend = ConvertToHtml(Head=HTMLExport[1], Body="<H2>Hydra Deployment Results</H2> " + HTMLExport[0]).OutString()
    fSendMail(mailserver, mailfrom, mailto, mailreplyto, "Hydra Deployment Results", ToSend, True)
    OutputDataGrid.ClearSelection()

# Enable or disable the "Start" button depending on different criterias
GroupFound = False
for row in OutputDataGrid.Rows:
    if row.Index == OutputDataGrid.RowCount - 1:
        continue
    if row.Cells[0].Tag.GroupID != "0":
        GroupFound = True
        break
OutputDataGrid.Columns[9].Visible = GroupFound
ExportGroupMenu.Enabled = GroupFound
SelectAllGroupMenu.Enabled = GroupFound
DeSelectAllGroupMenu.Enabled = GroupFound

CheckboxesFound = nbCheckedBoxes > 0
if not CheckboxesFound:
    ActionButton.Enabled = False
    return

SequenceAssigned = False
if not SequenceLoaded:
    for row in OutputDataGrid.Rows:
        if row.Cells[7].Value and row.Cells[0].Tag.GroupID != "0":
            SequenceAssigned = True
            break

if not SequenceLoaded and not SequenceAssigned:
    ActionButton.Enabled = False
    return

ReadyToInstall = False
for row in OutputDataGrid.Rows:
    if row.Cells[7].Value and row.Cells[4].Value <= 0:
        ReadyToInstall = True
        break
ActionButton.Enabled = ReadyToInstall

# Parse the grid for objects ready to start the selected Sequence
if SequencesTreeView.SelectedNode.Parent == None:
    return

SelectedRows = [row for row in OutputDataGridTab[GridIndex].Rows if row.Cells[7].Value and row.Cells[0].Tag.GroupID == "0" and (row.Cells[8].Value == 0 or row.Cells[4].Value < 0)]
VariablesReloaded = False
if len(SelectedRows) == 0:
    return

if SecurityCode != "":
    CodePrompt = Read_SecurityCode(SecurityCode, SequenceName)
    if CodePrompt != "OK":
        return "err"
elif DisplayWarning:
    ReallyDeploy = Read_StartSequence(SequenceName)
    if ReallyDeploy != "OK":
        return "err"

if len(SequencePanelVariable) == 0:
    VariablesQuery = Set_CurrentSequenceToObjectsVariables()
    
    
    
    
    
    
    
    
    
    
    
    
    
    
else:
    # Variables already set: ask if they should be reused or reloaded
    ReloadVariables = MessageBox.Show("Variables have been already defined for " + SequenceName + "\r\nDo you want to reuse them ?", "WARNING", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
    if ReloadVariables == "no":
        for j in range(len(SequencePanelVariable)):
            SequencePanelVariable[j].Text = ""  # Clear the variables set in the Sequence Panel
        SequencePanelVariable = []
        VariablesQuery = Set_CurrentSequenceToObjectsVariables()  # Query and set the variables
        VariablesReloaded = True

if VariablesQuery == "error":
    # The variable query has been cancelled, the sequence start is stopped
    return "err"

for row in SelectedRows:
    # Set the status of the objects
    row.Cells[2].Value = "Sequence loaded: " + SequenceName
    row.DefaultCellStyle.BackColor = "White"
    row.Cells[0].Tag.GroupID = "0"  # Set the GroupId to 0: doesn't belong to any group

SchedulerTemp = 0
GroupInuse = (len(GroupsUsed) > 0) and (GroupsUsed != "0")  # Are some groups in use  # Assign the current sequence to the free and checked objects. If groups are in use, or the sequence has changed, assign a new ID too (2nd parameter)
Set_CurrentSequenceToObjects(SelectedRows, (SelectionChanged or GroupInuse or VariablesReloaded), False)

def Set_AssignSequenceToObjects(UseScheduler, GroupToAssign = False):
    # Assign the current sequence to objects in a group
    if len(SequencePanelVariable) == 0:
        # The current loaded sequence doesn't have any variable set: if there are some, they are kept
        VariablesQuery = Set_CurrentSequenceToObjectsVariables()  # Query and set the variables, if any

    if VariablesQuery == "error":
        # The variable query has been cancelled, the sequence start is stopped
        return

    NewGroup = False
    if GroupToAssign != False:
        # A group has been set via a the Sequence Tree right click
        OutputDataGrid.ClearSelection()
        for i in range(OutputDataGrid.RowCount - 1):
            if OutputDataGrid.Rows[i].Cells[0].Tag.GroupID == GroupToAssign:
                # Select the objects matching the group
                OutputDataGrid.Rows[i].Cells[0].Selected = True
    else:
        # Define a new group
        GroupToSet = Read_InputBoxDialog("Group", "Set the Group Name to assign '" + SequenceName + "':", "")
        if GroupToSet == "":
            return
        if GroupToSet in GroupsUsed:
            # Check if the name is already given to another group
            MessageBox.Show("The Group Name '" + GroupToSet + "' is already assigned.", "Group Name", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            return
        NewGroupName = GroupToSet
        GroupsUsed.append(NewGroupName)  # Extend GroupsUsed with the newly created group name
        NewGroup = True

    SelectedRows = [row for row in OutputDataGrid.Rows if row.Cells[0].Selected]  # Determine the objects selected
    if UseScheduler:
        # Ask for a Scheduler if needed
        Scheduler = Read_DateTimePicker("Enter the start for " + SequenceName)
        if Scheduler == "":
            return
        SchedulerTemp = Scheduler
    else:
        SchedulerTemp = 0  # No Scheduler set

    DataGridViewCellStyleBold = System.Windows.Forms.DataGridViewCellStyle()
    DataGridViewCellStyleBold.Alignment = 16
    DataGridViewCellStyleBold.Font = System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Bold, 3, 0)
    DataGridViewCellStyleBold.ForeColor = System.Drawing.Color.FromArgb(255, 0, 0, 0)
    DataGridViewCellStyleBold.SelectionBackColor = System.Drawing.Color.FromArgb(255, 51, 153, 255)
    DataGridViewCellStyleBold.SelectionForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)

    DataGridViewCellStyleRegular = System.Windows.Forms.DataGridViewCellStyle()
    DataGridViewCellStyleRegular.Alignment = 16
    DataGridViewCellStyleRegular.Font = System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Regular, 3, 0)
    DataGridViewCellStyleRegular.ForeColor = System.Drawing.Color.FromArgb(255, 0, 0, 0)
    DataGridViewCellStyleRegular.SelectionBackColor = System.Drawing.Color.FromArgb(255, 51, 153, 255)
    DataGridViewCellStyleRegular.SelectionForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)

    for row in SelectedRows:
        # Assign the current sequence to the selected objects
        row.Cells[2].Value = "Sequence assigned: " + SequenceName
        row.DefaultCellStyle.BackColor = "White"
        row.Cells[0].Style = DataGridViewCellStyleBold
        row.Cells[2].Style = DataGridViewCellStyleRegular
        row.Cells[0].ReadOnly = True
        if (DisplayWarning or (SecurityCode != "")):
            row.Cells[7].Value = (GrpCheckedOnWarning == "True")  # Auto uncheck the objects on warning if the option is set in the user's settings
        if NewGroup:
            # A new group has been define
            row.Cells[0].Tag.GroupID = NewGroupName  # Set the name of the group on the Cell's Tag
            row.Cells[9].Value = NewGroupName  # Set the group name in the Group Column
            if GrpShowThreads == "True":
                row.Cells[9].Value += " (" + MaxThreadsText.Text + ")"  # Add also the number of Threads if the option is set in the user's settings         
                   
    SequenceAssigned = True
if SchedulerTemp != 0:
    row.Cells[3].Value = SchedulerTemp.ToLongTimeString()
else:
    row.Cells[3].Value = "Pending"

Set-CurrentSequenceToObjects(SelectedRows, True, True)

if UseAScheduler:
    for row in SelectedRows:
        row.Cells[3].Value = SchedulerTemp.ToLongTimeString()

if len(SequencePanelVariable) != 0:
    for j in range(len(SequencePanelVariable)):
        SequencePanelVariable[j].Text = ""
    SequencePanelVariable = []

Set-ActionButtonState()
def CellValue(IndexOfGrid, Row, Id1, Id2, Id3, Id4, Id5, Id6, Id8):
    if Id1 != "#":
     OutputDataGridTab[IndexOfGrid].Rows[Row].Cells[1].Value = Id1
    if Id2 != "#":
        OutputDataGridTab[IndexOfGrid].Rows[Row].Cells[2].Value = Id2
    if Id3 != "#":
        OutputDataGridTab[IndexOfGrid].Rows[Row].Cells[3].Value = Id3
    if Id4 != "#":
        OutputDataGridTab[IndexOfGrid].Rows[Row].Cells[4].Value = Id4
    if Id5 != "#":
        OutputDataGridTab[IndexOfGrid].Rows[Row].Cells[5].Value = Id5
    if Id6 != "#":
        OutputDataGridTab[IndexOfGrid].Rows[Row].Cells[6].Value = Id6
    if Id8 != "#":
        OutputDataGridTab[IndexOfGrid].Rows[Row].Cells[8].Value = Id8
def CheckAll(Check):
    for i in range(OutputDataGrid.RowCount - 1):
        OutputDataGrid.Rows[i].Cells[7].Value = Check
OutputDataGrid.RefreshEdit()
Get-CountCheckboxes()
def CloseForm():
    if ResetSettings == True:
        return
PosRegistry = True
WindowPositions = [Form.Top, Form.Left, Form.Size.Width, Form.Size.Height, SplitContainer1.Size.Width, SplitContainer1.Size.Height, 
        SplitContainer1.SplitterDistance, SplitContainer2.Size.Width, SplitContainer2.Size.Height, SplitContainer2.SplitterDistance]
for Pos in WindowPositions:
        if Pos < -20 or Pos > 3000:
            PosRegistry = False
if PosRegistry:
        # Save all variable values
        pass
if ShowSearchBox == "True":
    SequencesTreeViewTopPosition = 20
else:
    SequencesTreeViewTopPosition = 0
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosFormY" -Value Form.Top -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosFormX" -Value Form.Left -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosFormW" -Value Form.Width - 2 * FormBorderSize -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosFormH" -Value Form.Height - 2 * FormBorderSize - FormHeaderSize -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosSplit1W" -Value SplitContainer1.Size.Width - 2 * FormBorderSize - 15 -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosSplit1H" -Value SplitContainer1.Size.Height -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosSplit1D" -Value SplitContainer1.SplitterDistance -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosSplit2W" -Value SplitContainer2.Size.Width - 15 -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosSplit2H" -Value SplitContainer2.Size.Height -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosSplit2D" -Value SplitContainer2.SplitterDistance -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqTreeTop" -Value SequencesTreeView.Top - SequencesTreeViewTopPosition -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqTreeLeft" -Value SequencesTreeView.Left -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqTreeWidth" -Value SequencesTreeView.Width -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqTreeHeight" -Value SequencesTreeView.Height + SequencesTreeViewTopPosition -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqPanelTop" -Value SequenceTasksPanel.Top -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqPanelLeft" -Value SequenceTasksPanel.Left -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqPanelWidth" -Value SequenceTasksPanel.Width -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqPanelHeight" -Value SequenceTasksPanel.Height -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PanelBottomTop" -Value PanelBottom.Top -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PanelBottomWidth" -Value PanelBottom.Width - 2 * FormBorderSize - 15 -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "DataGridTabControlWidth" -Value DataGridTabControl.Width - 2 * FormBorderSize - 15 -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "DataGridTabControlHeight" -Value DataGridTabControl.Height -Type String -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "LastDirObjects" -Value LastDirObjects -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "LastDirSequences" -Value LastDirSequences -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "LastDirExportGroup" -Value LastDirExportGroup -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "LastDirImportGroup" -Value LastDirImportGroup -Force
#Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "WelcomeScreen" -Value "False" -Force

def CopyObjectsToTab(Move, Tab):
    if Tab == -1:
        Set-NewTab()
        TabToCopy = TabPageIndex
    else:
        TabToCopy = Tab

SelectedObjects = OutputDataGrid.SelectedCells.select(ExpandProperty="Value")
CurrentTab = DataGridTabControl.SelectedTab.TabIndex
OutputDataGrid = OutputDataGridTab[TabToCopy]
Add-ObjectListToGrid(SelectedObjects, "")
OutputDataGrid = OutputDataGridTab[CurrentTab]
if Move == True:
    Set-RightClick_SetNewSelectionFromGrid(False)
def set_current_sequence_to_objects(objects_selected, new_id, belongs_to_group):
    # Assign the sequence to the selected objects
    if len(objects_selected) == 0:
        return
    
    if new_id:
        # A new Sequence ID has to be used: increase all arrays and set the respective attributes
        if use_a_scheduler:
            scheduler_temp = scheduler_var
        
        new_seq = {}
        new_seq["ScriptBlock"] = [script_block_loaded]
        new_seq["ScriptBlockComment"] = [script_block_comment_loaded]
        new_seq["ScriptBlockPreLoad"] = False
        new_seq["ScriptBlockVariable"] = [variables_query]
        new_seq["ScriptBlockModule"] = [sequence_import_module_loaded]
        new_seq["ScriptBlockCheckboxes"] = []
        new_seq["SequenceScheduler"] = scheduler_temp
        new_seq["SequenceSchedulerExpired"] = False
        new_seq["SequenceSendMail"] = send_mail
        new_seq["MaxCheckedObjects"] = max_objects
        new_seq["SequenceLabel"] = sequence_name
        new_seq["MaxThreads"] = max_threads_text.text
        new_seq["SecurityCode"] = security_code
        new_seq["BelongsToGroup"] = belongs_to_group
        
        if display_warning:
            new_seq["DisplayWarning"] = 1
        else:
            new_seq["DisplayWarning"] = 0
        
        sequences.append(new_seq)
        output_data_grid_sequence.append([])
        runspace_collection.append([])
        runspace_pool.append([])
        job_nb.append([])
        concurrent_jobs.append([])
        sequence_tab_index += output_data_grid.tag.tab_page_index
        output_data_grid_sequence[len(sequences) - 1] = []
    
    rows_in_use = [row.index for row in output_data_grid_sequence[len(sequences) - 1]]
    
    for row in objects_selected:
        # Set or reset the cells values
        row.cells[4].value = -1  # Set the Sequence Step to -1, Pending
        row.cells[8].value = len(sequences) - 1  # Assign the last Sequence ID to the object
        
        if row.index not in rows_in_use:
            # A new object has been added to the Sequence
            output_data_grid_sequence[len(sequences) - 1].append(row)  # Add the row to the output_data_grid_sequence associated to the Sequence ID
    
    # Recreate the ScriptBlockCheckboxes array to avoid that a pointer is created (Reference Type): solves the issue with multiple start of a Sequence with Preload
    sequences[len(sequences) - 1]["ScriptBlockCheckboxes"] = []
    
    for i in range(len(sequence_panel_checkbox)):
        checkbox = CheckBox()
        checkbox.fore_color = sequence_panel_checkbox[i].fore_color
        checkbox.location = Size(sequence_panel_checkbox[i].left, sequence_panel_checkbox[i].top)
        checkbox.checked = sequence_panel_checkbox[i].checked
        checkbox.auto_size = True
        checkbox.maximum_size = Size(500, 15)
        checkbox.font = sequence_panel_checkbox[i].font
        checkbox.text = sequence_panel_checkbox[i].text
        
        sequences[len(sequences) - 1]["ScriptBlockCheckboxes"].append(checkbox)
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
def Set_CurrentSequenceToObjectsVariables():
    if len(SequencePanelVariable) != 0:
        return
    VarP = 0
    UseAScheduler = False
    for i in range(1, nbVariableTypes + 1):
        SeqVariableHash[i] = {}
    for i in range(len(SeqVariablesPos)):
        TypePos = VariableTypes.index(SeqVariablesPos[i].Type.lower())
        SeqVariables = SeqVariablesPos[i].Value
        VariableQuery = eval(VariableCommand[TypePos])
        if VariableQuery == "":
            messagebox.show("  Process cancelled  ", VariableTypes[TypePos], MessageBoxButtons.OK, MessageBoxIcon.Stop)
            for j in range(len(SequencePanelVariable)):
                SequencePanelVariable[j].Text = ""
            SequencePanelVariable = []
            return "error"
        VariableName = SeqVariablesPos[i].Name
        if VariableTypes[TypePos] != "secretinputbox":
            Set_SequencePanelVariable("  {}: {}".format(VariableName, VariableQuery), VarP)
        else:
            Set_SequencePanelVariable("  {}: ********".format(VariableName), VarP)
        VarP += 1
        SeqVariableHash[TypePos].add(VariableName, VariableQuery)
        if TypePos == 7:
            UseAScheduler = True
            SchedulerVar = VariableQuery
            
            
            
  # Return the whole hash created
def SeqVariableHash():
    return

# Create all variables needed and set their default values
CSVSeparator = ";"
CSVTempPath = "C:\Temp\HydraExport.csv"
XLSXTempPath = "C:\Temp\HydraExport.xlsx"
HTMLTempPath = "C:\Temp\HydraExport.html"
LogFilePath = "C:\Temp\Hydra.log"
CentralLogPath = "C:\Temp"
LogFileEnabled = "True"
Colors = {"OK": "#FF90EE90", "BREAK": "#FFADD8E6", "STOP": "#FFF08080", "CANCELLED": "#FFC0C0C0"}
DefaultThreads = "10"
DisplayWarning = False
SCCM_ConfigMgrSiteServer = ""
SCCM_SiteCode = ""
NoSplashScreen = "False"
WelcomeScreen = "True"
DebugMode = 0
ColorBackground = "#FF61B598"
ColorSequences = "#FFFFFFFF"
ColorSequencesRunning = "#FFFFFFE6"
LastDirObjects = ""
LastDirSequences = ""
LastDirExportGroup = ""
LastDirImportGroup = ""
ShowSearchBox = "True"
FileLoaded = "False"
CountriesList = "$HydraSettingsPath\Hydra_Countries.sccm"
ADQueriesList = "$HydraSettingsPath\Hydra_ADQueries.qry"
SequencesListPath = "$HydraSettingsPath\Hydra_Sequences.lst"
nbCheckedBoxes = 0
LoadedFiles = ""
SequenceListExpanded = "True"
RowHeaderVisible = "False"
CheckBoxesKeepState = "False"
EMailSMTPServer = ""
EMailSendFrom = ""
EMailSendTo = ""
EMailReplyTo = ""
GrpCheckedOnWarning = "False"
GrpShowThreads = "True"
PosFormX = 100
PosFormY = 10
PosFormW = 1150
PosFormH = 750
PosSplit1W = 1084
PosSplit1H = 688
PosSplit1D = 230
PosSplit2W = 215
PosSplit2H = 688
PosSplit2D = 350
SeqTreeTop = 35
SeqTreeLeft = 10
SeqTreeWidth = 215
SeqTreeHeight = 310
SeqPanelTop = 25
SeqPanelLeft = 10
SeqPanelWidth = 215
SeqPanelHeight = 290
PanelBottomTop = 600
PanelBottomWidth = 840
DataGridTabControlWidth = 825
DataGridTabControlHeight = 580
TabLook = "1"
TimerIntervalDefault = 1000
TimerInterval = TimerIntervalDefault
SendMail = False
TimerSet = False
NewGroupName = 0
RemovedFromSeqID = 0
SelectionChanged = False
VariablesQuery = None

TabColorPalette = {"DodgerBlue": "LightBlue", "Crimson": "LightCoral", "SpringGreen": "LightGreen", "Yellow": "Khaki", "DarkGray": "LightGray", "DarkOrchid": "Orchid", "Sienna": "Chocolate"}

Sequences = []












RunspaceCollection = []
RunspacePool = []
SequencesToParse = []
OutputDataGridSequence = []
JobNb = []
ConcurrentJobs = []
SequenceAssigned = False
SequenceTabIndex = []
GroupsUsed = []
GroupsRunning = []

def Set_GroupMaxThreads():
    NewMaxThread = int(input("Max. Threads\nEnter the new maximum threads for the Group: "))
    if NewMaxThread < 1 or NewMaxThread > 1000:
        return
    SelectedRows = [row for row in OutputDataGrid.Rows[0:OutputDataGrid.Rows.Count - 1] if row.Cells[0].Tag.GroupID == GroupsFoundForMaxThreads]
    for row in SelectedRows:
        Sequences[row.Cells[8].Value].MaxThreads = NewMaxThread
        row.Cells[9].Value = str(GroupsFoundForMaxThreads)
        if GrpShowThreads == "True":
            row.Cells[9].Value += " (" + str(NewMaxThread) + ")"

def Set_ObjectsState():
    ObjectsLabel.Text = "Total Objects: " + str(OutputDataGrid.RowCount - 1) + ", Selected: " + str(nbCheckedBoxes)

def Set_ObjectsToGroup(GroupToAdd):
    RowRef = None
    for i in range(OutputDataGrid.RowCount):
        if OutputDataGrid.Rows[i].Cells[0].Tag.GroupID == GroupToAdd:
            RowRef = OutputDataGrid.Rows[i]
            break
    if RowRef is None:
        return
    SeqValue = RowRef.Cells[8].Value
    MaxThreadValue = Sequences[RowRef.Cells[8].Value].MaxThreads
    SeqName = Sequences[SeqValue].SequenceLabel
    DataGridViewCellStyleBold = System.Windows.Forms.DataGridViewCellStyle()
    DataGridViewCellStyleBold.Alignment = 16
    DataGridViewCellStyleBold.Font = System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Bold, 3, 0)
    DataGridViewCellStyleBold.ForeColor = System.Drawing.Color.FromArgb(255, 0, 0, 0)
    DataGridViewCellStyleBold.SelectionBackColor = System.Drawing.Color.FromArgb(255, 51, 153, 255)
    DataGridViewCellStyleBold.SelectionForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
    DataGridViewCellStyleRegular = System.Windows.Forms.DataGridViewCellStyle()
    DataGridViewCellStyleRegular.Alignment = 16
    DataGridViewCellStyleRegular.Font = System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Regular, 3, 0)
    DataGridViewCellStyleRegular.ForeColor = System.Drawing.Color.FromArgb(255, 0, 0, 0)
    DataGridViewCellStyleRegular.SelectionBackColor = System.Drawing.Color.FromArgb(255, 51, 153, 255)
    DataGridViewCellStyleRegular.SelectionForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
    for item in OutputDataGrid.SelectedCells:
        RowIndex = item.RowIndex
        Set_CellValue(GridIndex, RowIndex, 0, "Sequence assigned: " + SeqName, "Pending", 0, "#", "#", SeqValue)
        OutputDataGrid.Rows[RowIndex].Cells[0].ReadOnly = True
        OutputDataGrid.Rows[RowIndex].Cells[0].Tag.GroupID = GroupToAdd
        OutputDataGrid.Rows[RowIndex].DefaultCellStyle.BackColor = "White"
        OutputDataGrid.Rows[RowIndex].Cells[0].Style = DataGridViewCellStyleBold
        OutputDataGrid.Rows[RowIndex].Cells[2].Style = DataGridViewCellStyleRegular
        OutputDataGrid.Rows[RowIndex].Cells[9].Value = str(GroupToAdd)
        if GrpShowThreads == "True":
            OutputDataGrid.Rows[RowIndex].Cells[9].Value += " (" + str(MaxThreadValue) + ")"
        OutputDataGridSequence[SeqValue].append(OutputDataGrid.Rows[RowIndex])
        
        
        
        
        
        
        
        
        
        
def Set_RecreateGroups():
    # Parse the grid and recreate the groups after deleting operations have been performed
    # Search for all Sequence ID's assigned to groups in the grid
    SeqFound = []
    for i in range(DataGridTabControl.TabCount):
        TabGridID = DataGridTabControl.TabPages[i].Tag.TabPageIndex  # Get the Grid ID
        SeqFound += [row.Cells[8].Value for row in OutputDataGridTab[TabGridID].Rows if row.Cells[0].Tag.GroupID != "0"]
    if len(SeqFound) == 0:
        # No Sequence ID relative to any group found: no group defined (anymore)
        GroupsUsed = []
        return
    for Seq in SeqFound:
        OutputDataGridSequence[Seq] = []
    # Reset all OutputDataGridSequence relative to the Sequence ID's found
    for i in range(OutputDataGrid.RowCount - 1):
        # Parse the grid and refill the OutputDataGridSequence arrays
        if OutputDataGrid.Rows[i].Cells[0].Tag.GroupID != "0":
            SeqId = OutputDataGrid.Rows[i].Cells[8].Value
            OutputDataGridSequence[SeqId].append(OutputDataGrid.Rows[i])
    GroupsUsed = []
    # Parse all tabs, get the name of all the Groups and store them in GroupsUsed
    for i in range(DataGridTabControl.TabCount):
        TabGridID = DataGridTabControl.TabPages[i].Tag.TabPageIndex  # Get the Grid ID
        GroupsUsed += [cell.Tag.GroupID for cell in OutputDataGridTab[TabGridID].Rows.Cells if cell.ColumnIndex == 0 and cell.Tag.GroupID != 0]
    GroupsUsed = list(set(GroupsUsed))

def Set_ReloadSequenceList():
    # Reload the Sequence List for potential changes or correction
    SequencesTreeView.Nodes.Clear()
    SequencesTreeView_GetSequenceList(True)
    if FormSettingsSequenceExpandedRadioButton.Checked:
        # Collapse or expand depending on the user's settings
        SequencesTreeView.ExpandAll()
    else:
        SequencesTreeView.CollapseAll()

def Set_RightClick_CheckObject(Check):
    # Check or uncheck the objects selected
    for RowIndex in OutputDataGrid.SelectedCells.RowIndex:
        OutputDataGrid.Rows[RowIndex].Cells[7].Value = Check
    Get_CountCheckboxes()

def Set_RightClick_RemoveSelectionFromFiles(FileList, DeleteRows):
    # Remove the selected objects from the files they were loaded from
    IsRunningSeqSelected = len([row for row in OutputDataGrid.SelectedRows if row.Cells[4].Value > 0])
    if IsRunningSeqSelected > 0:
        return  # If some objects are running, exits
    for File in FileList:
        # Generate the objects to remove for each file
        ObjectsRemaining = []
        for SelectedItem in OutputDataGrid.SelectedCells:
            if OutputDataGrid.Rows[SelectedItem.RowIndex].Cells[6].Value == File:
                # The file name matches: adds the object in the array
                ObjectsRemaining.append(SelectedItem.Value)
        NewText = [line for line in Select_String(Path=File, Pattern=ObjectsRemaining, NotMatch=True).Line]
        # Recreate the list of objects removing the ones stored in the array
        Set_Content(Path=File, Content=NewText)
        # Recreate the objects file with the new content
    if DeleteRows:
        # Delete the objects from the grid too if needed
        Set_RightClick_SetNewSelectionFromGrid(False)

def Set_RightClick_SetNewSelectionFromGrid(Action):
    # Recreate the grid with or without the selected objects
    IsRunningSeqSelected = len([row for row in OutputDataGrid.SelectedRows if row.Cells[4].Value > 0])
    if IsRunningSeqSelected > 0:
        return  # If some objects are running, exits
    NewSelection = []
    NewSelectionCellStyle = []
    for i in range(OutputDataGrid.RowCount - 1):
        # Create a new selection based on Action
        if OutputDataGrid.Rows[i].Cells[0].Selected == Action:
            # If Action=True, add the selected objects, if Action=False, add the non-selected objects
            NewSelection.append(OutputDataGrid.Rows[i])
            NewSelectionCellStyle.append(OutputDataGrid.Rows[i].Cells[2].Style)
    SeqFound = [row.Cells[8].Value for row in NewSelection]
    SeqFound = list(set(SeqFound))
    for Seq in SeqFound:
        OutputDataGridSequence[Seq] = []
    # Reset the OutputDataGridSequence for all sequences found in the new selection
    OutputDataGrid.Rows.Clear()
    DataGridViewCellStyleBold = System.Windows.Forms.DataGridViewCellStyle()
    DataGridViewCellStyleBold.Alignment = 16
    DataGridViewCellStyleBold.Font = System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Bold, 3, 0)
    DataGridViewCellStyleBold.ForeColor = System.Drawing.Color.FromArgb(255, 0, 0, 0)
    DataGridViewCellStyleBold.SelectionBackColor = System.Drawing.Color.FromArgb(255, 51, 153, 255)
    DataGridViewCellStyleBold.SelectionForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
    for i in range(len(NewSelection)):
        # Recreate the new grid based on the newly created selection
        OutputDataGrid.Rows.Add(NewSelection[i].Cells.Value)
        # Clone the row and add the Tags
        OutputDataGrid.Rows[i].Cells[0].Tag = NewSelection[i].Cells[0].Tag
        if OutputDataGrid.Rows[i].Cells[0].Tag.GroupID != "0":
            # Group found
            OutputDataGrid.Rows[i].Cells[0].Style = DataGridViewCellStyleBold
            OutputDataGrid.Rows[i].Cells[0].ReadOnly = True
        SeqId = OutputDataGrid.Rows[i].Cells[8].Value
        OutputDataGridSequence[SeqId].append(OutputDataGrid.Rows[i])
        # Add the row to the OutputDataGridSequence of the sequence ID
        if OutputDataGrid.Rows[i].Cells[4].Value < 0:
            # A sequence ran already: don't reset the color
            OutputDataGrid.Rows[i].Cells[2].Style = NewSelectionCellStyle[i]
            OutputDataGrid.Rows[i].DefaultCellStyle.BackColor = OutputDataGrid.Rows[i].Cells[5].Value
            
                  
        
        
        
        
                
def Set_RecreateGroups():
    Get_CountCheckboxes()

def Set_RightClick_SetNewSelectionFromState(State, DeleteRows):
    IsRunningSeqSelected = len([row for row in OutputDataGrid.SelectedRows if row.Cells[4].Value > 0])
    if IsRunningSeqSelected > 0:
        return
    OutputDataGrid.ClearSelection()
    for i in range(OutputDataGrid.RowCount - 1):
        if OutputDataGrid.Rows[i].Cells[3].Value == State:
            OutputDataGrid.Rows[i].Cells[0].Selected = True
    if DeleteRows:
        Set_RightClick_SetNewSelectionFromGrid(False)
    else:
        Set_RecreateGroups()
        Get_CountCheckboxes()

def Set_RightClick_ShowProtocol():
    NewText = []
    ProtocolFound = False
    RowsToKeep = [OutputDataGrid.Rows[cell.RowIndex] for cell in OutputDataGrid.SelectedCells]
    for Row in RowsToKeep:
        NewText.append(f"{Row.Cells[0].Tag.StepProtocol}\r\n")
        if Row.Cells[0].Tag.StepProtocol is not None:
            ProtocolFound = True
    if not ProtocolFound:
        NewText = "No protocol found"
    OutputDataGridContextMenuObjectProtocol.DropDownItems[0].Text = NewText

def Set_RunspaceToObjects():
    RowChecked = [cell.RowIndex for cell in OutputDataGrid.Rows.Cells if cell.ColumnIndex == 7 and cell.Value == True]
    RowsSeqIDs = [OutputDataGrid.Rows[row].Cells[8] for row in RowChecked]
    SequenceIndex = [index.Value for index in RowsSeqIDs if index.Value != 0]
    for Index in SequenceIndex:
        if Sequences[Index].BelongsToGroup:
            if Sequences[Index].SecurityCode != "":
                CodePrompt = Read_SecurityCode(Sequences[Index].SecurityCode, Sequences[Index].SequenceLabel)
                if CodePrompt != "OK":
                    SequenceIndex = [i for i in SequenceIndex if i != Index]
                    continue
            elif Sequences[Index].DisplayWarning != False:
                ReallyDeploy = Read_StartSequence(Sequences[Index].SequenceLabel)
                if ReallyDeploy != "OK":
                    SequenceIndex = [i for i in SequenceIndex if i != Index]
                    continue
    if SequenceIndex is None:
        return
    MissingIndex = [index for index in SequenceIndex if index not in SequencesToParse]
    CentralLogFilePath = FormSettingsPathsValue[4] + "\\" + Environment.UserName + "\\"
    for Index in MissingIndex:
        if Sequences[Index].SequenceSchedulerExpired:
            for cell in OutputDataGridSequence[Index].Cells:
                if cell.ColumnIndex == 7:
                    cell.Value = False
            for cell in OutputDataGridSequence[Index].Cells:
                if cell.ColumnIndex == 3:
                    cell.Value = "Timer Expired"
            OutputDataGrid.RefreshEdit()
            continue
        if Sequences[Index].SequenceScheduler != 0:
            TimeDiff = datetime.now() - Sequences[Index].SequenceScheduler
            TimeDiffFormated = "{:02d}:{:02d}:{:02d}".format(TimeDiff.hours, TimeDiff.minutes, TimeDiff.seconds)
            if TimeDiff.total_seconds() <= 1:
                Sequences[Index].SequenceScheduler = 0
                Sequences[Index].SequenceSchedulerExpired = True
                for cell in OutputDataGridSequence[Index].Cells:
                    if cell.ColumnIndex == 7:
                        cell.Value = False
                for cell in OutputDataGridSequence[Index].Cells:
                    if cell.ColumnIndex == 3:
                        cell.Value = "Timer Expired"
                OutputDataGrid.RefreshEdit()
                continue
               
RunspaceCollection[Index] = []
SessionState = System.Management.Automation.Runspaces.InitialSessionState.CreateDefault()
SessionState.Variables.Add(System.Management.Automation.Runspaces.SessionStateVariableEntry("MyScriptInvocation", MyScriptInvocation, None))
SessionState.Variables.Add(System.Management.Automation.Runspaces.SessionStateVariableEntry("SequencePath", SequencePath, None))
SessionState.Variables.Add(System.Management.Automation.Runspaces.SessionStateVariableEntry("SequenceFullPath", SequenceAbsolutePath, None))
SessionState.Variables.Add(System.Management.Automation.Runspaces.SessionStateVariableEntry("CentralLogPath", CentralLogFilePath, None))

for Module in Sequences[Index].ScriptBlockModule:
    if Module.Type == "ImportPSSnapIn":
        SessionState.ImportPSSnapIn(Module.Name, None)
    if Module.Type == "ImportPSModulesFromPath":
        SessionState.ImportPSModulesFromPath(Module.Name)
    if Module.Type == "ImportPSModule":
        SessionState.ImportPSModule(Module.Name)

for i in range(1, nbVariableTypes+1):
    for key in Sequences[Index].ScriptBlockVariable[i].keys():
        SessionState.Variables.Add(System.Management.Automation.Runspaces.SessionStateVariableEntry(key, Sequences[Index].ScriptBlockVariable[i][key], None))

RunspacePool[Index] = RunspaceFactory.CreateRunspacePool(1, Sequences[Index].MaxThreads, SessionState, Host)
RunspacePool[Index].Open()

if Index not in SequencesToParse:
    SequencesToParse.append(Index)

JobNb[Index] = 0
ConcurrentJobs[Index] = 1

def Set_SearchBox():
    global SequencesTreeViewTopPosition
    if ShowSearchBox == "True":
        SequencesTreeViewTopPosition = 20
    else:
        SequencesTreeViewTopPosition = 0
    SequencesTreeView.Top = int(SeqTreeTop) + SequencesTreeViewTopPosition
    SequencesTreeView.Height = SeqTreeHeight - SequencesTreeViewTopPosition
    SearchTreeTextBox.Visible = (ShowSearchBox == "True")

def Set_SelectGroup(Group, Select):
    OutputDataGrid.ClearSelection()
    if Group == "All Groups":
        GroupsToSelect = [cell.Tag.GroupID for cell in OutputDataGrid.Rows.Cells if cell.ColumnIndex == 0 and cell.Tag.GroupID != 0]
        GroupsToSelect = list(set(GroupsToSelect))
        GroupsToSelect.sort()
    else:
        GroupsToSelect = [Group]
    
    for row in OutputDataGrid.Rows:
        if row.Cells[0].Tag.GroupID in GroupsToSelect:
            row.Cells[7].Value = Select
    
    OutputDataGrid.EndEdit()
    Get_CountCheckboxes()

def Set_SequenceFinished():
    Timer.Enabled = False
    SequenceRunning = False
    ObjectsLabel.Text = ""
    Send_MailLog()
    Write_Log()
    for Item in SequencesToParse:
        for Row in OutputDataGridSequence[Item]:
            Row.Cells[7].ReadOnly = False
    Reset_Runspaces()
    Reset_SequenceArrays()
    Get_CountCheckboxes()
    MenuMain.Items["Cancel"].Enabled = False
    MenuToolStrip.Items["Cancel All"].Enabled = False
    MenuMain.Items.DropDown.Items["Reset All Objects"].Enabled = True
    MenuMain.Items.DropDown.Items["Clear Grid"].Enabled = True

def Set_SequencePanelCheckbox(Text, SeqPos):
    SequencePanelTempCheckbox = System.Windows.Forms.Checkbox()
    Pos_Y = 43 * SeqPos + 45
    SequencePanelTempCheckbox.Location = System.Drawing.Size(15, Pos_Y)
    SequencePanelTempCheckbox.Checked = True
    SequencePanelTempCheckbox.ForeColor = 'Black'
    SequencePanelTempCheckbox.AutoSize = True
    SequencePanelTempCheckbox.MaximumSize = System.Drawing.Size(500, 15)
    SequencePanelTempCheckbox.Font = Drawing.Font("Microsoft Sans Serif", 8, Drawing.FontStyle.Regular)
    SequencePanelTempCheckbox.ForeColor = Drawing.Color.Black
    SequencePanelTempCheckbox.Text = Text
    SequencePanelCheckbox.append(SequencePanelTempCheckbox)
    SequenceTasksPanel.Controls.Add(SequencePanelCheckbox[SeqPos])
    
    
    
    
    
    
def Set_SequencePanelLabel(Text, Style, Color, SeqPos):
    SequencePanelTempLabel = System.Windows.Forms.Label()
    SequencePanelTempLabel.Text = Text
    Pos_Y = 43 * SeqPos + 62
    SequencePanelTempLabel.Location = System.Drawing.Size(15, Pos_Y)
    SequencePanelTempLabel.AutoSize = True
    SequencePanelTempLabel.MaximumSize = System.Drawing.Size(500, 15)
    SequencePanelTempLabel.Font = Drawing.Font("Microsoft Sans Serif", 8, Drawing.FontStyle[Style])
    SequencePanelTempLabel.ForeColor = Drawing.Color[Color]
    SequencePanelLabel.append(SequencePanelTempLabel)
    SequenceTasksPanel.Controls.Add(SequencePanelLabel[SeqPos])

def Set_SequencePanelTitle(Text, Color):
    SequencePanelTitleLabel = System.Windows.Forms.Label()
    SequencePanelTitleLabel.Text = Text
    SequencePanelTitleLabel.Location = System.Drawing.Size(10, 10)
    SequencePanelTitleLabel.AutoSize = True
    SequencePanelTitleLabel.MaximumSize = System.Drawing.Size(500, 15)
    SequencePanelTitleLabel.Font = Drawing.Font("Tahoma", 9, Drawing.FontStyle.Underline)
    SequencePanelTitleLabel.ForeColor = Drawing.Color[Color]
    SequenceTasksPanel.Controls.Add(SequencePanelTitleLabel)

def Set_SequencePanelVariable(Text, VarPos):
    SequencePanelTempVariable = System.Windows.Forms.Label()
    SequencePanelTempVariable.Text = Text.replace("\n", "|")
    Pos_Y = 44 * MaxSteps + 16 * VarPos + 48
    SequencePanelTempVariable.Location = System.Drawing.Size(15, Pos_Y)
    SequencePanelTempVariable.AutoSize = True
    SequencePanelTempVariable.MaximumSize = System.Drawing.Size(500, 15)
    SequencePanelTempVariable.Font = Drawing.Font("Microsoft Sans Serif", 8, Drawing.FontStyle.Regular)
    SequencePanelTempVariable.ForeColor = Drawing.Color.Blue
    SequencePanelVariable.append(SequencePanelTempVariable)
    SequenceTasksPanel.Controls.Add(SequencePanelVariable[VarPos])

def Set_SettingsSubMenu(SubMenu):
    FormSettingsPathsGroupBox.Visible = False
    FormSettingsColorsGroupBox.Visible = False
    FormSettingsColorsGUIGroupBox.Visible = False
    FormSettingsSequenceSearchGroupBox.Visible = False
    FormSettingsMiscGroupBox.Visible = False
    FormSettingsSequenceGroupBox.Visible = False
    FormSettingsRowHeaderGroupBox.Visible = False
    FormSettingsCheckBoxesGroupBox.Visible = False
    FormSettingsMailGroupBox.Visible = False
    FormSettingsGroupsWarningGroupBox.Visible = False
    FormSettingsGroupsThreadsGroupBox.Visible = False
    FormSettingsTabsLookGroupBox.Visible = False

    if SubMenu == "Paths":
        FormSettingsPathsGroupBox.Visible = True
    elif SubMenu == "Colors":
        FormSettingsColorsGroupBox.Visible = True
        FormSettingsColorsGUIGroupBox.Visible = True
    elif SubMenu == "Misc":
        FormSettingsSequenceSearchGroupBox.Visible = True
        FormSettingsMiscGroupBox.Visible = True
        FormSettingsSequenceGroupBox.Visible = True
        FormSettingsRowHeaderGroupBox.Visible = True
        FormSettingsCheckBoxesGroupBox.Visible = True
    elif SubMenu == "Mail":
        FormSettingsMailGroupBox.Visible = True
    elif SubMenu == "Groups":
        FormSettingsGroupsWarningGroupBox.Visible = True
        FormSettingsGroupsThreadsGroupBox.Visible = True
    elif SubMenu == "Tabs":
        FormSettingsTabsLookGroupBox.Visible = True

def Set_StartupSettings(Invocation):
    if Host.Version.Major < 4:
        System.Reflection.Assembly.LoadWithPartialName("System.Windows.Forms")
        System.Windows.Forms.MessageBox.Show("You need Powershell 4 or higher to run this version of Hydra.\n\nPlease install a newer version", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        exit
        
        
        
        
        
        
        
        
if not os.path.exists('HKCU:\Software\Hydra'):
    # Create the Hydra registry structure if it's missing
    os.makedirs('HKCU:\Software\Hydra', exist_ok=True)
else:
    if not os.path.exists('HKCU:\Software\Hydra\3'):
        # Clean-up the Hydra3 registry
        shutil.copytree('HKCU:\Software\Hydra', 'HKCU:\Software\Hydra3')
        shutil.rmtree('HKCU:\Software\Hydra', ignore_errors=True)
        os.makedirs('HKCU:\Software\Hydra', exist_ok=True)
        shutil.copytree('HKCU:\Software\Hydra3', 'HKCU:\Software\Hydra\3', dirs_exist_ok=True)
        shutil.rmtree('HKCU:\Software\Hydra3', ignore_errors=True)

if not os.path.exists('HKCU:\Software\Hydra\5'):
    # Copy the settings of Hydra4 to Hydra5
    if os.path.exists('HKCU:\Software\Hydra4'):
        shutil.copytree('HKCU:\Software\Hydra4', 'HKCU:\Software\Hydra\5', dirs_exist_ok=True)
        to_delete = ["DataGridHeight", "DataGridWidth", "PanelBottomTop", "PanelBottomWidth", "PosFormH", "PosFormW", "PosFormX", "PosFormY", "PosSplit1D", "PosSplit1H", "PosSplit1W", "PosSplit2D", "PosSplit2H", "PosSplit2W", "SeqPanelHeight", "SeqPanelLeft", "SeqPanelTop", "SeqPanelWidth", "SeqTreeHeight", "SeqTreeLeft", "SeqTreeTop", "SeqTreeWidth", "WelcomeScreen"]
        for item in to_delete:
            # Remove obsolete Hydra4 variables
            try:
                del os.environ['HKCU:\Software\Hydra\5'][item]
            except KeyError:
                pass
    else:
        os.makedirs('HKCU:\Software\Hydra\5', exist_ok=True)

# Define global variables
PSScriptName = os.path.basename(__file__)
HydraBinPath = os.path.dirname(sys.argv[0])
MyScriptInvocation = sys.argv
HydraSettingsPath = os.path.join(os.path.dirname(__file__), 'Settings')
HydraGUIPath = os.path.join(os.path.dirname(__file__), 'GUI')
HydraDocsPath = os.path.join(os.path.dirname(__file__), 'Docs')
SequenceName = ""
ResetSettings = False
SequenceRunning = False

SetDefaultSettings()  # Set the default variables settings
GetRegistrySettings()  # Get registry values and replace the default ones defined the step before

if SequencesListParam is not None:
    # Hydra has been started with a Sequences List as parameter
    SequencesListPath = SequencesListParam  # Set the global variable SequencesListPath with this parameter

os.chdir(os.path.dirname(__file__))  # Set the current directory to the Hydra path

if not os.path.exists(SequencesListPath):
    # No Sequences List found: exit
    ctypes.windll.user32.MessageBoxW(0, "Unable to find the Sequences List", "Error", 0x10)
    sys.exit()

# Load the Sequence Variables Types
SetStartupSettings_SeqVariables()  # Show or Hide the Powershell console

class Window(ctypes.Structure):
    _fields_ = [("handle", ctypes.c_void_p)]

consolePtr = ctypes.windll.kernel32.GetConsoleWindow()
ctypes.windll.user32.ShowWindow(consolePtr, DebugMode)  # 0 to make the Powershell console invisible, 5 to make the Powershell console visible

def SetStartupSettings_SeqVariables():
    # Define the variables types
    nbVariableTypes = 10
    SeqVariableHash = [0] * (nbVariableTypes + 1)
    VariableCommand = [0] * (nbVariableTypes + 1)
    VariableTypes = ["", "inputbox", "multilineinputbox", "selectfile", "selectfolder", "combobox", "multicheckbox", "scheduler", "credentials", "secretinputbox", "credentialbox"]

    # Create a script block based on a command to any of the variable types. It will pass SeqVariables defined by the user
    VariableCommand[1] = '''
        if len(SeqVariables.split(';')) == 0:
            Read_InputBoxDialog()
        elif len(SeqVariables.split(';')) == 1:
            Read_InputBoxDialog(SeqVariables.split(';')[0])
        elif len(SeqVariables.split(';')) == 2:
            Read_InputBoxDialog(SeqVariables.split(';')[0], SeqVariables.split(';')[1])
        elif len(SeqVariables.split(';')) >= 3:
            Read_InputBoxDialog(SeqVariables.split(';')[0], SeqVariables.split(';')[1], SeqVariables.split(';')[2])
    '''      
                       
        
 
 
 
 
 
 
 
 
 
 
VariableCommand[2] = '''
    if len(SeqVariables.split(';')) == 0:
        Read_MultiLineInputBoxDialog()
    elif len(SeqVariables.split(';')) == 1:
        Read_MultiLineInputBoxDialog(SeqVariables.split(';')[0])
    elif len(SeqVariables.split(';')) == 2:
        Read_MultiLineInputBoxDialog(SeqVariables.split(';')[0], SeqVariables.split(';')[1])
    elif len(SeqVariables.split(';')) >= 3:
        Read_MultiLineInputBoxDialog(SeqVariables.split(';')[0], SeqVariables.split(';')[1], SeqVariables.split(';')[2])
'''

VariableCommand[3] = '''
    if len(SeqVariables.split(';')) == 0:
        Read_OpenFileDialog()
    elif len(SeqVariables.split(';')) == 1:
        Read_OpenFileDialog(SeqVariables.split(';')[0])
    elif len(SeqVariables.split(';')) == 2:
        Read_OpenFileDialog(SeqVariables.split(';')[0], SeqVariables.split(';')[1])
    elif len(SeqVariables.split(';')) >= 3:
        LastParamPos = len(SeqVariables.split(';')[0]) + len(SeqVariables.split(';')[1]) + 2
        Read_OpenFileDialog(SeqVariables.split(';')[0], SeqVariables.split(';')[1], SeqVariables[LastParamPos:])
'''

VariableCommand[4] = '''
    fBrowse_Folder_Modern(SeqVariables)
'''

VariableCommand[5] = '''
    if len(SeqVariables.split(';')) == 0:
        Read_ComboBoxDialog()
    elif len(SeqVariables.split(';')) == 1:
        Read_ComboBoxDialog(SeqVariables.split(';')[0])
    elif len(SeqVariables.split(';')) == 2:
        Read_ComboBoxDialog(SeqVariables.split(';')[0], SeqVariables.split(';')[1])
    elif len(SeqVariables.split(';')) >= 3:
        Read_ComboBoxDialog(SeqVariables.split(';')[0], SeqVariables.split(';')[1], SeqVariables.split(';')[2])
'''

VariableCommand[6] = '''
    if len(SeqVariables.split(';')) == 0:
        Read_MultiCheckboxList()
    elif len(SeqVariables.split(';')) == 1:
        Read_MultiCheckboxList(SeqVariables.split(';')[0])
    elif len(SeqVariables.split(';')) == 2:
        Read_MultiCheckboxList(SeqVariables.split(';')[0], SeqVariables.split(';')[1])
    elif len(SeqVariables.split(';')) >= 3:
        Read_MultiCheckboxList(SeqVariables.split(';')[0], SeqVariables.split(';')[1], SeqVariables.split(';')[2])
'''

VariableCommand[7] = '''
    Read_DateTimePicker(SeqVariables)
'''

VariableCommand[8] = '''
    Read_Credentials(SeqVariables)
'''

VariableCommand[9] = '''
    if len(SeqVariables.split(';')) == 0:
        Read_InputDialogBoxSecret()
    elif len(SeqVariables.split(';')) == 1:
        Read_InputDialogBoxSecret(SeqVariables.split(';')[0])
    elif len(SeqVariables.split(';')) >= 2:
        Read_InputDialogBoxSecret(SeqVariables.split(';')[0], SeqVariables.split(';')[1])
'''

VariableCommand[10] = '''
    if len(SeqVariables.split(';')) == 0:
        Read_Credentialbox()
    elif len(SeqVariables.split(';')) == 1:
        Read_Credentialbox(SeqVariables.split(';')[0])
    elif len(SeqVariables.split(';')) >= 2:
        Read_Credentialbox(SeqVariables.split(';')[0], SeqVariables.split(';')[1])
'''

def Set_TabColor(ColorSelected, ColorUnselected):
    DataGridTabControl.SelectedTab.Tag.ColorSelected = ColorSelected
    DataGridTabControl.SelectedTab.Tag.ColorUnSelected = ColorUnselected
    DataGridTabControl.Refresh()

def Set_TabStyle():
    if FormSettingsTabsLookCheckedRadioButton[0].Checked == True:
        DataGridTabControl.DrawMode = "Normal"
        TabLook = "0"
    if FormSettingsTabsLookCheckedRadioButton[1].Checked == True:
        DataGridTabControl.DrawMode = "Normal"
        DataGridTabControl.DrawMode = "OwnerDrawFixed"
        DataGridTabControl.Remove_DrawItem(DataGridTabControl_DrawItemHandlerColorsFull)
        DataGridTabControl.Remove_DrawItem(DataGridTabControl_DrawItemHandlerColorsLine)
        DataGridTabControl.Add_DrawItem(DataGridTabControl_DrawItemHandlerColorsFull)
        TabLook = "1"
    if FormSettingsTabsLookCheckedRadioButton[2].Checked == True:
        DataGridTabControl.DrawMode = "Normal"
        DataGridTabControl.DrawMode = "OwnerDrawFixed"
        DataGridTabControl.Remove_DrawItem(DataGridTabControl_DrawItemHandlerColorsFull)
        DataGridTabControl.Remove_DrawItem(DataGridTabControl_DrawItemHandlerColorsLine)
        DataGridTabControl.Add_DrawItem(DataGridTabControl_DrawItemHandlerColorsLine)
        TabLook = "2"
        
        
        
        
        
        
        
        
        
        
        
        
def Set_Timer():
    Timer.Stop()
    Form.Refresh()
    Timer.Interval = TimerInterval
    Timer.Start()
    Form.Refresh()

def Set_UnAssignSequenceToObjects():
    GroupName = OutputDataGrid.Rows[OutputDataGrid.SelectedCells[0].RowIndex].Cells[0].Tag.GroupID
    SeqId = OutputDataGrid.Rows[OutputDataGrid.SelectedCells[0].RowIndex].Cells[8].Value
    DataGridViewCellStyleRegular = System.Windows.Forms.DataGridViewCellStyle()
    DataGridViewCellStyleRegular.Alignment = 16
    DataGridViewCellStyleRegular.Font = System.Drawing.Font("Microsoft Sans Serif", 8.25, Drawing.FontStyle.Regular, 3, 0)
    DataGridViewCellStyleRegular.ForeColor = System.Drawing.Color.FromArgb(255, 0, 0, 0)
    DataGridViewCellStyleRegular.SelectionBackColor = System.Drawing.Color.FromArgb(255, 51, 153, 255)
    DataGridViewCellStyleRegular.SelectionForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
    
    for SelectedItem in OutputDataGrid.SelectedCells:
        OutputDataGrid.Rows[SelectedItem.RowIndex].Cells[7].ReadOnly = False
        OutputDataGrid.Rows[SelectedItem.RowIndex].Cells[0].Tag.GroupID = "0"
        OutputDataGrid.Rows[SelectedItem.RowIndex].Cells[0].Style = DataGridViewCellStyleRegular
        OutputDataGrid.Rows[SelectedItem.RowIndex].Cells[9].Value = "-"
        if Sequences[OutputDataGrid.Rows[SelectedItem.RowIndex].Cells[8].Value].SequenceScheduler != 0:
            OutputDataGrid.Rows[SelectedItem.RowIndex].Cells[3].Value = "Pending"
            SelectionChanged = True
    
    SequencesTreeView.SelectedNode = SequencesTreeView.Nodes[0]
    Set_RecreateGroups()
    Get_CountCheckboxes()

def Set_UserSettings():
    for i in range(4):
        ColorHex = "#FF{:X2}{:X2}{:X2}".format(FormSettingsColorsButton[i].BackColor.R, FormSettingsColorsButton[i].BackColor.G, FormSettingsColorsButton[i].BackColor.B)
        Colors.Set_Item(FormSettingsColorsButton[i].Name, ColorHex)
        RegColorName = "Color_" + FormSettingsColorsButton[i].Name
        Set_ItemProperty(Path="HKCU:\Software\Hydra\5", Name=RegColorName, Value=ColorHex)
    
    for i in range(4):
        Set_Variable(Name=FormSettingsPathsVariable[i], Value=FormSettingsPathsText[i].Text, Scope="Script", Force=True)
        Set_ItemProperty(Path="HKCU:\Software\Hydra\5", Name=FormSettingsPathsVariable[i], Value=FormSettingsPathsText[i].Text)
    
    FormSettingsMailText[2].Text = FormSettingsMailText[2].Text.replace(";", ",")
    for i in range(4):
        Set_Variable(Name=FormSettingsMailVariable[i], Value=FormSettingsMailText[i].Text, Scope="Script", Force=True)
        Set_ItemProperty(Path="HKCU:\Software\Hydra\5", Name=FormSettingsMailVariable[i], Value=FormSettingsMailText[i].Text)
    
    ColorHex = "#FF{:X2}{:X2}{:X2}".format(FormSettingsColorsGUIBackButton.BackColor.R, FormSettingsColorsGUIBackButton.BackColor.G, FormSettingsColorsGUIBackButton.BackColor.B)
    Set_Variable(Name="ColorBackground", Value=ColorHex, Scope="Script", Force=True)
    Set_ItemProperty(Path="HKCU:\Software\Hydra\5", Name="ColorBackground", Value=ColorHex)
    
    ColorHex = "#FF{:X2}{:X2}{:X2}".format(FormSettingsColorsGUISeqButton.BackColor.R, FormSettingsColorsGUISeqButton.BackColor.G, FormSettingsColorsGUISeqButton.BackColor.B)
    Set_Variable(Name="ColorSequences", Value=ColorHex, Scope="Script", Force=True)
    Set_ItemProperty(Path="HKCU:\Software\Hydra\5", Name="ColorSequences", Value=ColorHex)
    
    ColorHex = "#FF{:X2}{:X2}{:X2}".format(FormSettingsColorsGUISeqRunButton.BackColor.R, FormSettingsColorsGUISeqRunButton.BackColor.G, FormSettingsColorsGUISeqRunButton.BackColor.B)
    Set_Variable(Name="ColorSequencesRunning", Value=ColorHex, Scope="Script", Force=True)
    Set_ItemProperty(Path="HKCU:\Software\Hydra\5", Name="ColorSequencesRunning", Value=ColorHex)
    
    Form.BackColor = ColorBackground
    SequencesTreeView.BackColor = ColorSequences
    SequenceTasksPanel.BackColor = ColorSequences
    
    if FormSettingsSequenceShowSearchRadioButton.Checked:
        Set_ItemProperty(Path="HKCU:\Software\Hydra\5", Name="ShowSearchBox", Value="True", Force=True)
        Set_Variable(Name="ShowSearchBox", Value="True", Scope="Script", Force=True) 
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
    else:
        subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "ShowSearchBox", "-Value", "False", "-Force"])
        subprocess.run(["Set-Variable", "-Name", "ShowSearchBox", "-Value", "False", "-Scope", "Script", "-Force"])

if FormSettingsSplashScreenCheckBox.Checked:
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "NoSplashScreen", "-Value", "False", "-Force"])
else:
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "NoSplashScreen", "-Value", "True", "-Force"])

if FormSettingsDebugScreenCheckBox.Checked:
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "DebugMode", "-Value", "5", "-Force"])
else:
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "DebugMode", "-Value", "0", "-Force"])

if FormSettingsSequenceExpandedRadioButton.Checked:
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "SequenceListExpanded", "-Value", "True", "-Force"])
    SequencesTreeView.ExpandAll()
else:
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "SequenceListExpanded", "-Value", "False", "-Force"])
    SequencesTreeView.CollapseAll()

if FormSettingsRowHeaderGroupBoxVisibleRadioButton.Checked:
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "RowHeaderVisible", "-Value", "True", "-Force"])
else:
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "RowHeaderVisible", "-Value", "False", "-Force"])

OutputDataGrid.RowHeadersVisible = FormSettingsRowHeaderGroupBoxVisibleRadioButton.Checked

if FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked:
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "CheckBoxesKeepState", "-Value", "True", "-Force"])
else:
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "CheckBoxesKeepState", "-Value", "False", "-Force"])

if FormSettingsLogCheckBox.Checked:
    subprocess.run(["Set-Variable", "-Name", "LogFileEnabled", "-Value", "True", "-Scope", "Script", "-Force"])
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "LogFileEnabled", "-Value", "True", "-Force"])
else:
    subprocess.run(["Set-Variable", "-Name", "LogFileEnabled", "-Value", "False", "-Scope", "Script", "-Force"])
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "LogFileEnabled", "-Value", "False", "-Force"])

if FormSettingsGroupsWarningCheckedRadioButton.Checked:
    subprocess.run(["Set-Variable", "-Name", "GrpCheckedOnWarning", "-Value", "True", "-Scope", "Script", "-Force"])
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "GrpCheckedOnWarning", "-Value", "True", "-Force"])
else:
    subprocess.run(["Set-Variable", "-Name", "GrpCheckedOnWarning", "-Value", "False", "-Scope", "Script", "-Force"])
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "GrpCheckedOnWarning", "-Value", "False", "-Force"])

if FormSettingsGroupsThreadsVisibleRadioButton.Checked:
    subprocess.run(["Set-Variable", "-Name", "GrpShowThreads", "-Value", "True", "-Scope", "Script", "-Force"])
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "GrpShowThreads", "-Value", "True", "-Force"])
else:
    subprocess.run(["Set-Variable", "-Name", "GrpShowThreads", "-Value", "False", "-Scope", "Script", "-Force"])
    subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "GrpShowThreads", "-Value", "False", "-Force"])

Show-GroupThreads()
Set-SearchBox()
Set-TabStyle()
subprocess.run(["Set-ItemProperty", "-Path", "HKCU:\Software\Hydra\5", "-Name", "TabLook", "-Value", "$TabLook", "-Force"])
def View_ColumnsSizeAuto():
    OutputDataGrid.Columns[0].Width = 100
    OutputDataGrid.Columns[0].AutoSizeMode = 'Fill'
    OutputDataGrid.Columns[0].FillWeight = 50
    OutputDataGrid.Columns[2].Width = 150
    OutputDataGrid.Columns[2].AutoSizeMode = 'Fill'
    OutputDataGrid.Columns[2].FillWeight = 150
    OutputDataGrid.Columns[3].Width = 100
    OutputDataGrid.Columns[3].AutoSizeMode = 'None'
def View_ColumnsSizeManual():
    for i in range(4):
        OutputDataGrid.Columns[i].AutoSizeMode = 'None'
def Set_View_Wrap():
    if any(item.Text == "Wrap Text" and item.Checked for item in MenuMain.Items.DropDown.items):
        # Wrap mode set
        for item in MenuMain.Items.DropDown.items:
            if item.Text == "Wrap Text":
                item.Checked = False  # Deactivate the wrap mode
        OutputDataGrid.RowsDefaultCellStyle.WrapMode = 'False'
        OutputDataGrid.Columns[2].DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        OutputDataGrid.AutoSizeRowsMode = 'AllCellsExceptHeaders'
        OutputDataGrid.Refresh()
        OutputDataGrid.AutoSizeRowsMode = 'None'
    else:
        # Activate the wrap mode
        for item in MenuMain.Items.DropDown.items:
            if item.Text == "Wrap Text":
                item.Checked = True
        OutputDataGrid.RowsDefaultCellStyle.WrapMode = 'True'
        OutputDataGrid.Columns[2].DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft
        OutputDataGrid.AutoSizeRowsMode = 'AllCellsExceptHeaders'
        OutputDataGrid.Refresh()

def Show_GroupThreads():
    for row in OutputDataGrid.Rows:
        if row.Index == OutputDataGrid.RowCount - 1 or row.Cells[9].Value == "-":
            continue
        row.Cells[9].Value = row.Cells[0].Tag.GroupID  # Get the name of the Group, in Cells[0].Tag.GroupID
        if GrpShowThreads == "True":
            row.Cells[9].Value += " ({})".format(Sequences[row.Cells[8].Value].MaxThreads)  # Show the threads values

def Show_PickColor(Color):
    ColorDialog = System.Windows.Forms.ColorDialog()
    ColorDialog.AllowFullOpen = True
    ColorPicked = ColorDialog.ShowDialog()
  
    if ColorPicked == "Cancel":
        return "Cancel"
    else:
        HexColor = "#FF{:02X}{:02X}{:02X}".format(ColorDialog.Color.R, ColorDialog.Color.G, ColorDialog.Color.B)
        return HexColor

def Show_SequenceSteps():
    SeqId = OutputDataGrid.Rows[OutputDataGrid.SelectedCells[0].RowIndex].Cells[8].Value  # Get the sequence ID based on the 1st object selected
    SequenceTasksPanel.Controls.Clear()
    SequencePanelTitleLabel = System.Windows.Forms.Label()
    SequencePanelTitleLabel.Text = Sequences[SeqId].SequenceLabel
    SequencePanelTitleLabel.Location = System.Drawing.Size(10, 10)
    SequencePanelTitleLabel.AutoSize = True
    SequencePanelTitleLabel.MaximumSize = System.Drawing.Size(500, 15)
    SequencePanelTitleLabel.Font = Drawing.Font("Tahoma", 9, Drawing.FontStyle.Underline)
    SequencePanelTitleLabel.ForeColor = Drawing.Color.DarkViolet
    SequenceTasksPanel.Controls.Add(SequencePanelTitleLabel)
    for SeqPosition in range(len(Sequences[SeqId].ScriptBlockComment)):
        SequencePanelTempCheckbox = Sequences[SeqId].ScriptBlockCheckboxes[SeqPosition]
        Sequences[SeqId].ScriptBlockCheckboxes[SeqPosition] = System.Windows.Forms.CheckBox()
        Sequences[SeqId].ScriptBlockCheckboxes[SeqPosition].ForeColor = Drawing.Color.DarkViolet
        Sequences[SeqId].ScriptBlockCheckboxes[SeqPosition].Location = System.Drawing.Size(SequencePanelTempCheckbox.Left, SequencePanelTempCheckbox.Top)
        Sequences[SeqId].ScriptBlockCheckboxes[SeqPosition].Checked = SequencePanelTempCheckbox.Checked
        Sequences[SeqId].ScriptBlockCheckboxes[SeqPosition].AutoSize = True
        Sequences[SeqId].ScriptBlockCheckboxes[SeqPosition].MaximumSize = System.Drawing.Size(500, 15)
        Sequences[SeqId].ScriptBlockCheckboxes[SeqPosition].Font = Drawing.Font("Microsoft Sans Serif", 8, Drawing.FontStyle.Regular)
        Sequences[SeqId].ScriptBlockCheckboxes[SeqPosition].Text = SequencePanelTempCheckbox.Text
        SequenceTasksPanel.Controls.Add(Sequences[SeqId].ScriptBlockCheckboxes[SeqPosition])
        SequencePanelTempLabel = System.Windows.Forms.Label()
        SequencePanelTempLabel.Text = "  {}\n\n".format(Sequences[SeqId].ScriptBlockComment[SeqPosition])
        Pos_Y = 43 * SeqPosition + 62
        SequencePanelTempLabel.Location = System.Drawing.Size(15, Pos_Y)
        SequencePanelTempLabel.AutoSize = True
        SequencePanelTempLabel.MaximumSize = System.Drawing.Size(500, 15)
        SequencePanelTempLabel.Font = Drawing.Font("Microsoft Sans Serif", 8, Drawing.FontStyle.Italic)
        SequencePanelTempLabel.ForeColor = Drawing.Color.DarkViolet
        SequenceTasksPanel.Controls.Add(SequencePanelTempLabel)
        
        
        
        
        
        
        
        
        
VarPos = 0
for i in range(1, nbVariableTypes+1):
    for key in Sequences[SeqId].ScriptBlockVariable[i].keys():
        VarName = key
        VarValue = Sequences[SeqId].ScriptBlockVariable[i][key]
        SequencePanelTempVariable = System.Windows.Forms.Label()
        SequencePanelTempVariable.Text = f"{VarName} : {VarValue}".replace("\n", "|")
        Pos_Y = 44 * len(Sequences[SeqId].ScriptBlockComment) + 16 * VarPos + 48
        SequencePanelTempVariable.Location = System.Drawing.Size(15, Pos_Y)
        SequencePanelTempVariable.AutoSize = True
        SequencePanelTempVariable.MaximumSize = System.Drawing.Size(500, 15)
        SequencePanelTempVariable.Font = Drawing.Font("Microsoft Sans Serif", 8, Drawing.FontStyle.Italic)
        SequencePanelTempVariable.ForeColor = Drawing.Color.Blue
        SequenceTasksPanel.Controls.Add(SequencePanelTempVariable)
        VarPos += 1

def Start_Sequence():
    if Set_AssignSequenceToFreeObjects() == "err":
        return
    Set_RunspaceToObjects()
    SelectionChanged = False
    OutputDataGrid.Focus()
    for Item in SequencesToParse:
        CheckedObjects = sum(1 for cell in OutputDataGridSequence[Item].Cells if cell.ColumnIndex == 7 and cell.Value == True)
        if CheckedObjects > Sequences[Item].MaxCheckedObjects and Sequences[Item].MaxCheckedObjects != 0:
            System.Windows.Forms.MessageBox.Show(f"Too much objects selected for '{Sequences[Item].SequenceLabel}'\n\nMaximum allowed: {Sequences[Item].MaxCheckedObjects}, Selected: {CheckedObjects}", Sequences[Item].SequenceLabel, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop)
            for cell in OutputDataGridSequence[Item].Cells:
                if cell.ColumnIndex == 7:
                    cell.Value = False
            OutputDataGrid.RefreshEdit()
        OutputDataGridSequence_DataGridViewCellStyle = System.Windows.Forms.DataGridViewCellStyle()
        OutputDataGridSequence_DataGridViewCellStyle.Alignment = 16
        OutputDataGridSequence_DataGridViewCellStyle.Font = System.Drawing.Font("Microsoft Sans Serif", 8.25, 0, 3, 0)
        OutputDataGridSequence_DataGridViewCellStyle.ForeColor = System.Drawing.Color.FromArgb(255, 0, 0, 0)
        OutputDataGridSequence_DataGridViewCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(255, 51, 153, 255)
        OutputDataGridSequence_DataGridViewCellStyle.SelectionForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        for Row in OutputDataGridSequence[Item]:
            if Row.Cells[7].Value and Row.Cells[4].Value <= 0 and Row.Cells[4].Value != -5:
                Row.Cells[0].ReadOnly = True
                Row.Cells[7].ReadOnly = True
                Row.DefaultCellStyle.BackColor = ColorSequencesRunning
                Row.Cells[0].Tag.StepProtocol = []
                Row.Cells[2].Style = OutputDataGridSequence_DataGridViewCellStyle
                Row.Cells[3].Value = "Pending"
                Row.Cells[4].Value = 0
                Row.Cells[0].Tag.PreviousStateComment = ""
                Row.Cells[0].Tag.SharedVariable = None
    ActionButton.Enabled = False
    for item in MenuToolStrip.Items:
        if item.ToolTipText == "Clear the Grid":
            item.Enabled = False
    if not SequenceRunning:
        SequenceRunning = True
        ObjectsDone = 0
        Timer.Enabled = True
        for item in MenuMain.Items:
            if item.Text == "Cancel":
                item.Enabled = True
        for item in MenuToolStrip.Items:
            if item.ToolTipText == "Cancel All":
                item.Enabled = True
        for item in MenuMain.Items.DropDown.Items:
            if item.Text == "Reset All Objects":
                item.Enabled = False
        for item in MenuMain.Items.DropDown.Items:
            if item.Text == "Clear Grid":
                item.Enabled = False
        Set_Timer()
    if not TimerSet:
        Timer.Add_Tick(GetData)
        TimerSet = True

def Write_DebugReceiveOutput(ReceiveOutput):
    for i in range(len(ReceiveOutput)):
        if i == 0:
            print(f"Value 1 - Status: {ReceiveOutput[i]}")
        elif i == 1:
            print(f"Value 2 - Result state: {ReceiveOutput[i]}")
        elif i == 2:
            print(f"Value 3 - Color: {ReceiveOutput[i]}")
        elif i == 3:
            print(f"Value 4 - Shared value: {ReceiveOutput[i]}")
        elif i > 3:
            print(f"Value {i+1} - Error: {ReceiveOutput[i]}")

def Write_Log():
    if not LogFileEnabled:
        return
    ToLog = ""
    for Item in SequencesToParse:
        ToLog += f"{Sequences[Item].SequenceLabel}  -  {datetime.now().strftime('%m/%d/%Y')}\r\n"
        for Row in OutputDataGridSequence[Item]:
            if Row.Cells[0].Tag.StepProtocol is not None:
                ToLog += f"{' ; '.join(Row.Cells[0].Tag.StepProtocol).strip()} \r\n"
        ToLog += "\r\n\r\n"     
        
        
        
        
        
        import sys
import os
import clr

def Add_Content(Value, Path):
    with open(Path, "a") as file:
        file.write(Value)

HydraVersion = "6.00"
sys.StrictMode = True
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("PresentationFramework")

import System.Windows.Forms as WinForms
import System.Drawing as Drawing

WinForms.Application.EnableVisualStyles()
exec(open(HydraGUIPath + "\\Hydra5_Res.ps1").read())
exec(open(HydraGUIPath + "\\Hydra5_Form.ps1").read())
exec(open(HydraGUIPath + "\\Dialogs.ps1").read())
Load_Logo()

if NoSplashScreen != "True":
    FormSplashScreen.ShowDialog()

Add_Form()
Form.ShowDialog()
Form.Dispose()
Timer.Dispose()                                              