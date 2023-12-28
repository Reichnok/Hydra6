#######################################################################
#                                                                     #
#  Hydra 5 (IronSnap)                                                 #
#                                                                     #
#                                                                     #
# Version 5.56 (05.05.2020)                                           #
# ING Technology Services addition                                    #
#     Cierra Master Tasks                                             #
#     Restore VDA Checkpoint                                          #
#     Perform Sigle Checkpoint                                        #
#                                                               MAT   #
#######################################################################
#######################################################################

#                                                                     #

#  Hydra 5 (IronSnap)                                                 #

#                                                                     #

#                                                                     #

# Version 5.56 (18.04.2021)                                           #

#     Add Collection To Install collection in SCCM                    #

#     Add Collection To Install collection in SCCM                    #

#     FIX Cierramaster by asking for credentials for elevated actions #

#                                                               MAT   #

#######################################################################

 

 

   $SequencesListParam = ".\settings\Hydra_ING.lst"
   Param(  # Optional script parameter defining an alternative Sequence List

  [Parameter(Mandatory = $true)] $SequencesListParam

)

# ---------------------------------


function Add-ObjectListToGrid($ObjectList, $FilePath) {

  # Set the default Cells Values of the objects passed as argument

  foreach ($Obj in $ObjectList) {
    $Obj = ($Obj.Trim())
    if ($obj.Length -eq 0) { continue }  # Eliminate empty objects
    $RowID = $OutputDataGrid.Rows.Add($Obj, "0", "Pending", "Pending", 0, "", $FilePath, $True, 0, "-")
    $OutputDataGrid.Rows[$RowID].Cells[7].ReadOnly = $False
    $OutputDataGrid.Rows[$RowID].DefaultCellStyle.BackColor = "White"
    $ObjectOptions = New-Object -TypeName PSObject
    $ObjectOptions | Add-Member –MemberType NoteProperty –Name GroupID –Value "0"
    $ObjectOptions | Add-Member –MemberType NoteProperty –Name PreviousStateComment –Value ""
    $ObjectOptions | Add-Member –MemberType NoteProperty –Name StepProtocol –Value $Null
    $ObjectOptions | Add-Member –MemberType NoteProperty –Name SharedVariable –Value $Null
    $OutputDataGrid.Rows[$RowID].Cells[0].Tag = $ObjectOptions
  }

  Get-CountCheckboxes

}


function Cancel-AllForce {

  # Cancel all the sequences and release the grid without waiting for the Runspaces return status

  $IDsToCancel = @(foreach ($RowIndex in $RowsSelected) { $OutputDataGrid.Rows[$RowIndex].Cells[8].Value } ) | select -Unique

  foreach ($RowIndex in @(0..$($OutputDataGrid.RowCount - 2))) {
    # Set the Cells Values to "Cancel", re-enable the checkbox
    if (($OutputDataGrid.Rows[$RowIndex].Cells[4].Value -eq 0) -or ($OutputDataGrid.Rows[$RowIndex].Cells[4].Value -eq -5)) {
      # Sequence not started, or in Cancelling state
      $StateToSet = -6
      Set-CellValue $GridIndex $RowIndex "#" "Cancelled" "CANCELLED" $StateToSet $Colors.Get_Item("CANCELLED") "#" "#"
      $OutputDataGrid.Rows[$RowIndex].Cells[7].Value = $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked 
      $OutputDataGrid.Rows[$RowIndex].Cells[7].ReadOnly = $False
      $OutputDataGrid.Rows[$RowIndex].Cells[0].Tag.StepProtocol = @("`r`n$($OutputDataGrid.Rows[$RowIndex].Cells[0].Value)")
      $OutputDataGrid.Rows[$RowIndex].Cells[0].Tag.StepProtocol += @("   Cancelled - Not started")
      $OutputDataGrid.Rows[$RowIndex].DefaultCellStyle.BackColor = $Colors.Get_Item("CANCELLED")
      if ($OutputDataGrid.Rows[$RowIndex].Cells[0].Tag.GroupID -eq "0") { $OutputDataGrid.Rows[$RowIndex].Cells[0].ReadOnly = $False }
    }
    if ($OutputDataGrid.Rows[$RowIndex].Cells[4].Value -gt 0) {
      # Sequence started
      $StateToSet = -6
      Set-CellValue $GridIndex $RowIndex "#" "Cancelled" "CANCELLED" $StateToSet $Colors.Get_Item("CANCELLED") "#" "#"
      $OutputDataGrid.Rows[$RowIndex].Cells[7].Value = $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked 
      $OutputDataGrid.Rows[$RowIndex].Cells[7].ReadOnly = $False
      $OutputDataGrid.Rows[$RowIndex].Cells[0].Tag.StepProtocol = @("`r`n$($OutputDataGrid.Rows[$RowIndex].Cells[0].Value)")
      $OutputDataGrid.Rows[$RowIndex].Cells[0].Tag.StepProtocol += @("   Cancelled")
      $OutputDataGrid.Rows[$RowIndex].DefaultCellStyle.BackColor = $Colors.Get_Item("CANCELLED")
      if ($OutputDataGrid.Rows[$RowIndex].Cells[0].Tag.GroupID -eq "0") { $OutputDataGrid.Rows[$RowIndex].Cells[0].ReadOnly = $False }
    }
  }

  # Reset the schedulers, if any
  foreach ($ID in $IDsToCancel) {
    if ($Sequences[$ID].SequenceScheduler -ne 0) {
      # An object with a scheduler has been cancelled: all relative objects will be cancelled too
      $Script:Sequences[$ID].SequenceScheduler = 0
      $Script:Sequences[$ID].SequenceSchedulerExpired = $True
      foreach ($item in $OutputDataGridSequence[$ID]) {
        $StateToSet = -6
        Set-CellValue $GridIndex $($item.Index) "#" "Cancelled" "CANCELLED" $StateToSet $Colors.Get_Item("CANCELLED") "#" "#"
        $item.Cells[7].Value = $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked 
        $item.Cells[7].ReadOnly = $False
        $item.Cells[0].Tag.StepProtocol = @("`r`n$($OutputDataGrid.Rows[$RowIndex].Cells[0].Value)")
        $item.Cells[0].Tag.StepProtocol += @("   Cancelled - Not started")
        $item.DefaultCellStyle.BackColor = $Colors.Get_Item("CANCELLED")
      }
    }
  }

}


function Cancel-Sequence($RowsSelected) {

  # Cancel the Sequences of the objects passed as argument
  
  $IDsToCancel = @(foreach ($RowIndex in $RowsSelected) { $OutputDataGrid.Rows[$RowIndex].Cells[8].Value } ) | select -Unique

  foreach ($RowIndex in $RowsSelected) {
    # Set the Cells Values to "Cancel", re-enable the checkbox
    if ($OutputDataGrid.Rows[$RowIndex].Cells[4].Value -eq 0) {
      #Sequence not started
      $StateToSet = -6
      Set-CellValue $GridIndex $RowIndex "#" "Cancelled" "CANCELLED" $StateToSet $Colors.Get_Item("CANCELLED") "#" "#"
      $OutputDataGrid.Rows[$RowIndex].Cells[7].Value = $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked 
      $OutputDataGrid.Rows[$RowIndex].Cells[7].ReadOnly = $False
      $OutputDataGrid.Rows[$RowIndex].Cells[0].Tag.StepProtocol = @("`r`n$($OutputDataGrid.Rows[$RowIndex].Cells[0].Value)")
      $OutputDataGrid.Rows[$RowIndex].Cells[0].Tag.StepProtocol += @("   Cancelled - Not started")
      $OutputDataGrid.Rows[$RowIndex].DefaultCellStyle.BackColor = $Colors.Get_Item("CANCELLED")
    }
    if ($OutputDataGrid.Rows[$RowIndex].Cells[4].Value -gt 0) {
      #Sequence started
      $StateToSet = -5
      Set-CellValue $GridIndex $RowIndex "#" "Cancelling" "CANCELLING" $StateToSet $Colors.Get_Item("CANCELLED") "#" "#"
      $OutputDataGrid.Rows[$RowIndex].DefaultCellStyle.BackColor = $Colors.Get_Item("CANCELLED")
    }
  }

  # Reset the schedulers, if any
  foreach ($ID in $IDsToCancel) {
    if ($Sequences[$ID].SequenceScheduler -ne 0) {
      # An object with a scheduler has been cancelled: all relative objects will be cancelled too
      $Script:Sequences[$ID].SequenceScheduler = 0
      $Script:Sequences[$ID].SequenceSchedulerExpired = $True
      foreach ($item in $OutputDataGridSequence[$ID]) {
        $StateToSet = -6
        Set-CellValue $GridIndex $($item.Index) "#" "Cancelled" "CANCELLED" $StateToSet $Colors.Get_Item("CANCELLED") "#" "#"
        $item.Cells[7].Value = $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked 
        $item.Cells[7].ReadOnly = $False
        $item.Cells[0].Tag.StepProtocol = @("`r`n$($OutputDataGrid.Rows[$RowIndex].Cells[0].Value)")
        $item.Cells[0].Tag.StepProtocol += @("   Cancelled - Not started")
        $item.DefaultCellStyle.BackColor = $Colors.Get_Item("CANCELLED")
      }
    }
  }
  
}


function Clear-Grid {

  # Clear the grid, reset all arrays to their empty values

  $OutputDataGrid.Rows.Clear()
  $OutputDataGrid.Rows[0].Cells[7].Value = $False
  $OutputDataGrid.Rows[0].Cells[7].ReadOnly = $True
  $Script:nbCheckedBoxes = 0
  $ObjectsLabel.Text = ""
  $OutputDataGridContextMenuObject.Items[5].Text = ""
  $SequencesTreeView.SelectedNode = $SequencesTreeView.Nodes[0]

  Get-CountCheckboxes

}


$CreateRunspace = {

  param ($RunspaceScriptBlock, $RunspaceArg, $RSId, $SharedVariable)

  # Create a Runspace for the Object passed as $RunspaceArg, with the code $RunspaceScriptBlock
  # Use $RSId as Sequence Value and $SharedVariable for the variable shared between the successive steps

  $Powershell = [PowerShell]::Create().AddScript($RunspaceScriptBlock).AddArgument($RunspaceArg.Trim()).AddArgument($SharedVariable)
  $Powershell.RunspacePool = $RunspacePool[$RSId]
  $Script:RunspaceCollection[$RSId] += New-Object -TypeName PSObject -Property @{
    Runspace   = $PowerShell.BeginInvoke()
    PowerShell = $PowerShell  
  }

}


function Export-Result ($ExportFormat) {

  # Close the Export form and start the right Export, based on the format given as parameter

  $FormExport.Hide()
  switch ($ExportFormat) {
    0 { Export-ToCSV }
    1 { Export-ToExcel }
    2 { Export-ToHTML $True $True }
    3 { Send-Email }
  }

}


function Export-ToCSV {

  # Export and display the grid information in CSV format

  $Selection = @()
  if ($FormExportColCheckBox[1].Checked) { $Selection += 0 }
  if ($FormExportColCheckBox[2].Checked) { $Selection += 2 }
  if ($FormExportColCheckBox[3].Checked) { $Selection += 3 }

  New-Item $CSVTempPath -Type File -Force | Out-Null

  if ($FormExportHeaderCheckBox.Checked) {
    # Create the Header if necessary
    $ToPaste = (($OutputDataGrid.Columns[$Selection]) | select -ExpandProperty Name) -Join $CSVSeparator  # Create the header using the columns name
    if ($FormExportColCheckBox[4].Checked) {
      $ToPaste = ($ToPaste, "Sequence Name") -join $CSVSeparator
    }
    Add-Content -Path $CSVTempPath -Value $ToPaste  # Create the CSV file with the Header as content
  }

  if ($FormExportSelectionCheckBox.Checked -eq $True) {
    # Define the rows to export based on the export user's choice (Selected Objects)
    $OutputGridRows = $OutputDataGrid.Rows[ ($OutputDataGrid.SelectedCells | select -ExpandProperty RowIndex -Unique) ]
  }
  else {
    # (All Objects)
    $OutputGridRows = $OutputDataGrid.Rows[0..$($OutputDataGrid.RowCount - 2)]
  }

  foreach ($Row in $OutputGridRows) {
    # Get the information to export and add them in the CSV file
    $ToPaste = ($Row.Cells[$Selection] | select -ExpandProperty EditedFormattedValue) -Join $CSVSeparator
    if ($FormExportColCheckBox[4].Checked) {
      # Convert the SequenceID in its real name
      $ToPaste = ($ToPaste, $Sequences[$Row.Cells[8].EditedFormattedValue].SequenceLabel) -join $CSVSeparator
    }
    Add-Content -Path $CSVTempPath -Value $ToPaste 
  }

  Start-Process 'C:\windows\system32\notepad.exe' -ArgumentList $CSVTempPath

}


function Export-ToExcel {

  # Export and display the grid information in Excel

  Export-ToHTML $False $False  # Use the Exporet-ToHTML function to generate a temporary raw HTML file: create it without any style ($false) and don't open it ($false) 

  $Cc = [threading.thread]::CurrentThread.CurrentCulture  # Save the current regional settings
  [threading.thread]::CurrentThread.CurrentCulture = 'en-US'  # Set the Culture to en-US to avoid some bugs
  $Excel = New-Object -ComObject Excel.Application  # Open Excel and display the temporary HTML created file
  $Excel.Visible = $True
  $WorkBook = $Excel.Workbooks.Open($HTMLTempPath)
  $Excel.Windows.Item(1).Displaygridlines = $True
  $Now = (Get-Date).ToString("yyyyMMddHHssmm") 
  $NewXLSXName = (New-Object System.IO.FileInfo(Split-Path $XLSXTempPath -Leaf)).BaseName + "_" + $Now + (New-Object System.IO.FileInfo(Split-Path $XLSXTempPath -Leaf)).Extension
  $NewXLSXPath = Join-Path -Path (Split-Path $XLSXTempPath -Parent) -ChildPath $NewXLSXName
  $Workbook.SaveAs($NewXLSXPath)  # Save the newly created file
  [threading.thread]::CurrentThread.CurrentCulture = $Cc  # Set the regional settings back

}


function Export-CreateHTML($Object, $TaskResult, $Step, $SequenceName, $Color, $OnlySelection, $WithStyle) {

  # Create an HTML files based on the All or Selected objects, the columns to display as well as color and style

  $GridObjects = @()
  $ColumnSelection = @()

  # Create the Column Selection depending of the parameters booleans
  if ($Color) { $ColumnSelection += 'Color' }
  if ($Object) { $ColumnSelection += $OutputDataGrid.Columns[0].Name }
  if ($TaskResult) { $ColumnSelection += "Task*" }
  if ($Step) { $ColumnSelection += $OutputDataGrid.Columns[3].Name }
  if ($SequenceName) { $ColumnSelection += 'Sequence Name' }

  if ($OnlySelection) {
    # Define the Rows to export (Selected)
    $OutputGridRows = $OutputDataGrid.Rows[($OutputDataGrid.SelectedCells | select -ExpandProperty RowIndex -Unique)]
  }
  else {
    # (All)
    $OutputGridRows = $OutputDataGrid.Rows[0..$($OutputDataGrid.RowCount - 2)]
  }

  foreach ($Row in $OutputGridRows) {
    # For each of these Rows, create and add an objects with the cells values
    $Prop = [ordered]@{}
    if ($Object) { $Prop.Add($OutputDataGrid.Columns[0].Name, $Row.Cells[0].EditedFormattedValue) }
    if (($TaskResult) -and ($FormExportHeaderCheckBox.Checked)) { $Prop.Add($OutputDataGrid.Columns[2].Name, $Row.Cells[2].EditedFormattedValue) }
    if (($TaskResult) -and ($FormExportHeaderCheckBox.Checked -eq $False)) {
      # No Header to use: split the elements of Task Results and make a column for each of them
      $i = 0
      foreach ($item in $($Row.Cells[2].EditedFormattedValue -Split ";")) {
        $i++
        $Prop.Add("TaskResultPart $i", $item) 
      }
    }
    if ($Step) { $Prop.Add($OutputDataGrid.Columns[3].Name, $Row.Cells[3].EditedFormattedValue) }
    if ($SequenceName) { 
      if ($($Sequences[$Row.Cells[8].EditedFormattedValue].SequenceLabel) -ne $Null) {
        # If the Sequence has been assigned or already run, get its name from the object's Sequence ID
        $Prop.Add('Sequence Name', $Sequences[$Row.Cells[8].EditedFormattedValue].SequenceLabel) 
      }
      else {
        $Prop.Add('Sequence Name', " ") 
      }       
    } 
    if ($Color) {
      $ColorHex = $Row.Cells[5].EditedFormattedValue -replace "#FF", "#"
      $Prop.Add('Color', "###" + $ColorHex + "###")  # Add ### before and after the color HEX value: this will be removed in a following step
    }
    $obj = New-Object -Type PSObject -Property $Prop
    $GridObjects += $obj  # Add the object created in the array $GridObjects
  }

  if ($WithStyle) {
    # Create the HTML Style if needed
    $HTMLStyle = "<style>"
    $HTMLStyle += "body { background-color:#dddddd; font-family:Tahoma; font-size:12pt; }"
    $HTMLStyle += "td, th { border:1px solid black; border-collapse:collapse; }"  
    $HTMLStyle += "th { color:white; background-color:black; }"     
    $HTMLStyle += "table, tr, td, th { padding: 2px; margin: 0px }"      
    $HTMLStyle += "table { margin-left:50px; }"     
    $HTMLStyle += "</style>" 
  }
  else {
    $HTMLStyle = ""
  }

  $HTMLBody = $GridObjects | Select-Object $ColumnSelection | ConvertTo-Html -Fragment  # Create the HTML code filtering the columns
  if ($FormExportHeaderCheckBox.Checked -eq $False) {
    # The Header shouldn't be displayed
    $HTMLBody = $HTMLBody -replace "<tr><th>.*?</th></tr>", ""  # Suppress the Automatic Header
  }

  if ($Color) {
    # String manipulation to get the color at the correct place in the HTML code
    $HTMLBody = $HTMLBody -replace "><td>###", " bgcolor="
    $HTMLBody = $HTMLBody -replace "###</td>", ">"
    $HTMLBody = $HTMLBody -replace "<th>Color</th>", ""
  }

  return $HTMLBody, $HTMLStyle      

}


function Export-Group($GroupExported) {

  # Export Group(s) to a XML bundle

  if ($GroupExported -eq "All Groups") {
    # Set the GroupList file: it will contain links to the single XML Group files
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.InitialDirectory = $LastDirExportGroup
    $SaveFileDialog.Filter = "Hydra Group List (*.grouplist)|*.grouplist|All files|*.*" 
    $SaveFileDialog.ShowDialog() |  Out-Null
    $Script:LastDirExportGroup = Split-Path $SaveFileDialog.FileName
    if ($SaveFileDialog.FileName -eq "") { return }

    Set-Content -Path $SaveFileDialog.FileName -Value "# Hydra Group List"  # Create the file with a header
    $GroupsInCurrentGrid = @(($OutputDataGrid.Rows.Cells | where { ($_.ColumnIndex -eq 0) } | select -ExpandProperty Tag) | where { $_.GroupID -ne 0 } | select -ExpandProperty GroupID -Unique | Sort-Object)  # Enumerate the groups in the grid
    foreach ($GroupUsedItem in $GroupsInCurrentGrid) {
      # Loop in the list of Groups to export
      if ($GroupUsedItem -eq "0") { continue }  # Skip the value 0 that doesn't belong to any Group 
      $ExportFileName = Join-Path (Split-Path $SaveFileDialog.FileName) "$GroupUsedItem.group.xml"  # The name automatically generated are using the name of the Groups themselves
      Export-GroupSingle $GroupUsedItem $ExportFileName  # Call the Export-GroupSingle function to export $GroupUsedItem and save it into $ExportFileName
      Add-Content -Path $SaveFileDialog.FileName -Value "$GroupUsedItem.group.xml"  # Add the path of the newly created XML in the GroupList file
    }
  }
  else {
    # Export one Group only
    Export-GroupSingle $GroupExported ""  # Call the Export-GroupSingle function to export $GroupExported defined globally, the "" will display the Save window
  }

}


function Export-GroupSingle($GroupToExport, $FileForExport) {

  # Create an XML file with all attributes needed for a re-import

  $SeqId = 0
  $ObjectsList = @()
  for ($i = 0; $i -lt $OutputDataGrid.RowCount - 1; $i++) {  
    if ($OutputDataGrid.Rows[$i].Cells[0].Tag.GroupID -eq $GroupToExport) {
      # Extract the member of the group $GroupToExport
      $SeqID = $OutputDataGrid.Rows[$i].Cells[8].Value  # Get the corresponding Sequence ID
      $ObjectsList += $OutputDataGrid.Rows[$i].Cells[0].Value  # Create a list with the Group members
    }
  }

  if ($SeqID -eq 0) { return }  # SeqID=0, No group found

  if ($FileForExport -eq "") {
    # Get the name for saving if not given as parameter
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.InitialDirectory = $LastDirExportGroup
    $SaveFileDialog.Filter = "Hydra Group (*.group.xml)|*.group.xml|All files|*.*" 
    $SaveFileDialog.ShowDialog() |  Out-Null
    $Script:LastDirExportGroup = Split-Path $SaveFileDialog.FileName  # Save the last folder for registry settings saving on close
    if ($SaveFileDialog.FileName -eq "") { return }
    $FileForExport = $SaveFileDialog.FileName
  } 

  # Save the different attributes arrays in Export variables
  $ScriptBlockExport = $Sequences[$SeqID].ScriptBlock
  $ScriptBlockCommentExport = $Sequences[$SeqID].ScriptBlockComment
  $ScriptBlockVariableExport = $Sequences[$SeqID].ScriptBlockVariable
  $ScriptBlockModuleExport = $Sequences[$SeqID].ScriptBlockModule
  $ScriptBlockCheckboxesExport = $Sequences[$SeqID].ScriptBlockCheckboxes
  $ScriptBlockPreLoadExport = $Sequences[$SeqID].ScriptBlockPreLoad
  $SequenceSchedulerExport = $Sequences[$SeqID].SequenceScheduler
  $MaxThreadsExport = $Sequences[$SeqID].MaxThreads
  $MaxCheckedObjectsExport = $Sequences[$SeqID].MaxCheckedObjects
  $SequenceSendMailExport = $Sequences[$SeqID].SequenceSendMail
  $SequenceLabelExport = $Sequences[$SeqID].SequenceLabel
  $BelongsToGroupExport = $Sequences[$SeqID].BelongsToGroup
  $SequenceSchedulerExpiredExport = $Sequences[$SeqID].SequenceSchedulerExpired
  $SecurityCodeExport = $Sequences[$SeqID].SecurityCode
  $DisplayWarningExport = $Sequences[$SeqID].DisplayWarning
  $ExportVer = 1  # File Version number
  $GroupExported = $GroupToExport

  # Export the variables set to $FileForExport with the command Export-Clixml
  Get-Variable ExportVer, ObjectsList, GroupExported, ScriptBlockExport, ScriptBlockCommentExport, ScriptBlockPreLoadExport, ScriptBlockVariableExport, ScriptBlockModuleExport, ScriptBlockCheckboxesExport, MaxThreadsExport, SequenceSchedulerExport, MaxCheckedObjectsExport, SequenceSendMailExport, SequenceLabelExport, BelongsToGroupExport, SequenceSchedulerExpiredExport, SecurityCodeExport, DisplayWarningExport | Export-Clixml $FileForExport

}


function Export-Tabs {

  # Export all the Tabs in a file

  $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
  $SaveFileDialog.InitialDirectory = $LastDirExportGroup
  $SaveFileDialog.Filter = "Hydra Tabs (*.tabs)|*.tabs|All files|*.*" 
  $SaveFileDialog.ShowDialog() |  Out-Null
  $Script:LastDirExportGroup = Split-Path $SaveFileDialog.FileName
  if ($SaveFileDialog.FileName -eq "") { return }

  Set-Content -Path $SaveFileDialog.FileName -Value "# Hydra Tabs" # Create the header
  for ($i = 0; $i -lt $DataGridTabControl.TabCount; $i++) {
    # Loop in all Tabs
    $TabGridID = $DataGridTabControl.TabPages[$i].Tag.TabPageIndex  # Get the Grid ID
    $TabText = $DataGridTabControl.TabPages[$i].Text
    $TabColor1 = $DataGridTabControl.TabPages[$i].Tag.ColorSelected
    $TabColor2 = $DataGridTabControl.TabPages[$i].Tag.ColorUnSelected
    $TabSave = $TabText + ";" + $TabColor1 + ";" + $TabColor2 + ";" + $(($OutputDataGridTab[$TabGridID].Rows.Cells | where { $_.ColumnIndex -eq 0 } | select -ExpandProperty Value) -join ";")  # Get the objects names
    Add-Content -Path $SaveFileDialog.FileName -Value $TabSave
  }

}


function Export-ToHTML ($WithStyle, $OpenFile) {

  # Create an HTML file calling Export-CreateHTML and different parameters set by the user during the Export 

  $HTMLExport = Export-CreateHTML $FormExportColCheckBox[1].Checked $FormExportColCheckBox[2].Checked $FormExportColCheckBox[3].Checked $FormExportColCheckBox[4].Checked $FormExportColorCheckBox.Checked $FormExportSelectionCheckBox.Checked $WithStyle

  # Export the HTML created code to a complete HTML file 
  if ($WithStyle) {
    # If the Style is used, add a bigger header
    ConvertTo-Html -Head $($HTMLExport[1]) -Body "<H2>Sequence Results</H2> $($HTMLExport[0])" | Out-File $HTMLTempPath
  }
  else {
    ConvertTo-Html -Body "$($HTMLExport[0])" | Out-File $HTMLTempPath
  }

  if ($OpenFile) {
    # Automatically open the browser if needed
    Invoke-Expression $HTMLTempPath
  }

}


function Start-NewRunspaceScriptBlock($Row, $RSId) {

  # Define what to start and pass to the CreateRunspace function

  $Script:ConcurrentJobs[$RSId]++  # Increase the nunber of concurent jobs for the Sequence $RSId of one more

  if ($row.Cells[4].Value -eq 0) {
    # First step of the sequence for the object of $Row: create the protocol
    $row.Cells[0].Tag.StepProtocol += @("`r`n$($row.Cells[0].Value) - $($Sequences[$RSId].SequenceLabel) `r`n   Started at $CurrentTime")
  }

  if ($Sequences[$RSId].ScriptBlockCheckboxes[$($row.Cells[4].Value)].Checked) {
    # The step (checkbox) is enabled: start a new Runspace with the object corresponding ScriptBlock, Object Name, Sequence ID, Inter-Step Shared Variable
    & $CreateRunspace $Sequences[$RSId].ScriptBlock[$($row.Cells[4].Value)] $row.Cells[0].Value $RSId $row.Cells[0].Tag.SharedVariable
  }
  else {
    # The step is unchecked: create a runspace with a fake ScriptBlock returning "OK" and the old Cell content, Object Name, Sequence ID, Inter-Step Shared Variable
    & $CreateRunspace "return ""OK"", ""$($row.Cells[0].Tag.PreviousStateComment) (Step $($row.Cells[4].Value+1) Skipped)"" " $row.Cells[0].Value $RSId $row.Cells[0].Tag.SharedVariable
  }

  $row.Cells[4].Value++  # Increase the current step for the Row of 1
  $Script:JobNb[$RSId]++  # Increase the Job Number of 1 and set it to the Cell[1]: this matches the Runspace ID previously created

  # Set the Cells of $Row with the right paramters
  Set-CellValue $SequenceTabIndex[$RSId] $row.index $JobNb[$RSId] "Executing task:  $($Sequences[$RSId].ScriptBlockComment[$row.Cells[4].Value -1])" "Step $($row.Cells[4].Value)" $row.Cells[4].Value "#" "#" "#"
  
}


$GetData = { # Scriptblock run at every Timer Tick

  $CurrentTime = (Get-Date).ToLongTimeString()
  $StillRuning = $False  # Assume there is nothing running
  $TotalConcurrentJobs = 0  # Reset the number of concurrent running jobs to 0
  $Script:GroupsRunning = @()  # Empty the array of GroupsRunning

  foreach ($Item in $SequencesToParse) {
    # Loop into all Sequence ID's running to check the current states

    if ($Sequences[$Item].SequenceScheduler -ne 0) {
      # If a timer has expired, set the according SequenceScheduler to 0 to activate the start of the sequence
      $TimeDiff = New-TimeSpan $(Get-Date) $Sequences[$Item].SequenceScheduler
      $TimeDiffFormated = '{0:00}:{1:00}:{2:00}' -f $TimeDiff.Hours, $TimeDiff.Minutes, $TimeDiff.Seconds
      if ($TimeDiff.TotalSeconds -le 1) { 
        $Script:Sequences[$Item].SequenceScheduler = 0 
        $Script:Sequences[$Item].SequenceSchedulerExpired = $True
      }
    }

    foreach ($Row in $OutputDataGridSequence[$Item]) {
      # Loop into all objects of the Sequence $Item
    
      if ($row.Cells[7].Value -eq $False) { continue }  # Object unchecked: skip to the next one
      if ($Sequences[$Item].SequenceScheduler -ne 0) {
        # A scheduler is still running: actualize the countdown
        $Row.Cells[3].Value = $TimeDiffFormated
      }

      if (($row.Cells[4].Value -gt 0) -or ($row.Cells[4].Value -eq -5)) {
        # A sequence is running or is being cancelled
        Get-RowState $row $SequenceTabIndex[$Item]  # Check if the runspace has returned a value 
        if (($row.Cells[4].Value -gt 0) -or ($row.Cells[4].Value -eq -5)) {
          # The sequence is still running
          $StillRuning = $True
          if ($GroupsRunning -notcontains ($row.Cells[0].Tag.GroupID)) { $Script:GroupsRunning += $row.Cells[0].Tag.GroupID }  # Add the group ID to the array GroupsRunning
        }
      }
    }
  }

  foreach ($Item in $SequencesToParse) {
    # Loop into all Sequence ID's running to start new steps if necessary

    if ($Sequences[$Item].SequenceScheduler -ne 0) {
      # A timer is still running for this Sequence ID, skip to the next ID
      $StillRuning = $True
      continue
    }
    if ($ConcurrentJobs[$Item] -gt $RunspacePool[$Item].GetMaxRunspaces()) { continue }  # The number of concurrent runspaces for the ID has reached its maximum: skip to the next ID

    foreach ($Row in $OutputDataGridSequence[$Item]) {
      # Loop into all objects of the Sequence $Item

      if (($Item -ne $Row.Cells[8].Value) -or ($row.Cells[7].Value -eq $False)) { continue }  # Sequence ID mismatch or object unchecked: skip to the next one
      if (($row.Cells[4].Value -ge 0) -and ($row.Cells[2].Value -notlike "*Executing*")) {
        # The object is not executing anything, enter the block. If it is executing something, skip to the next row
        $StillRuning = $True
        if ($GroupsRunning -notcontains ($row.Cells[0].Tag.GroupID)) { $Script:GroupsRunning += $row.Cells[0].Tag.GroupID }  # Add the group ID to the array GroupsRunning

        if ($Row.Index -eq 0) {
          # Check the 1st line of the Sequence ID and check if a PreLoad task is running or has to run
          if (($Sequences[$Item].ScriptBlockCheckboxes[$($row.Cells[4].Value)].Text -like "*PreLoad*")) {
            # The step to execute is a PreLoad: set the ScriptBlockPreLoad of the Sequence ID to true and start the step for this object only
            $Script:Sequences[$Item].ScriptBlockPreLoad = $True
            Start-NewRunspaceScriptBlock $row $Item
            $Script:Sequences[$Item].ScriptBlockCheckboxes[$($row.Cells[4].Value - 1)].Checked = $False
          }
          elseif ($Sequences[$Item].ScriptBlockCheckboxes[$($row.Cells[4].Value)].Text -notlike "*PreLoad*") {
            # The step to execute is not a PreLoad: set the ScriptBlockPreLoad of the Sequence ID to false
            $Script:Sequences[$Item].ScriptBlockPreLoad = $False
          }
        }

        if ($Sequences[$Item].ScriptBlockPreLoad) { break }  # The object is currently running a PreLoad on the 1st object of the sequence ID: break the foreach
        Start-NewRunspaceScriptBlock $row $Item  # Start the next step for the current row

        if ($ConcurrentJobs[$Item] -gt $RunspacePool[$Item].GetMaxRunspaces()) { break }  # The number of concurrent runspaces for the ID has reached its maximum: skip to the next ID
      }

    }  

  }

  foreach ($Item in $SequencesToParse) { $TotalConcurrentJobs = $TotalConcurrentJobs + $ConcurrentJobs[$Item] - 1 }  # Count the number of all concurrent jobs of all sequences running

  $ObjectsLabel.Text = "Total Objects: $($OutputDataGrid.RowCount-1) ,   Running: $TotalConcurrentJobs ,   Done: $ObjectsDone"  # Display the current status of the running objects
  $RunningTask = @($OutputDataGrid.Rows | where { [int]($_.Cells[4].Value) -gt 0 }).Count  # Check if some objects are in a runing state and exits if any
  ($MenuToolStrip.Items | where { $_.ToolTipText -eq "Clear the Grid" }).Enabled = ($RunningTask -eq 0)

  if ($StillRuning -eq $False) {
    # Nothing is running anymore: the timer will be stopped, the runspaces and some variables will be cleared
    Set-SequenceFinished
  }

}


function Get-CountCheckboxes {

  # Count the number of objects checked, determine the state of the Start button and display the Objects state

  $Script:nbCheckedBoxes = 0
  $Script:nbCheckedBoxes = @($OutputDataGrid.Rows | where { $_.Cells[7].Value -eq $True }).Count

  Set-ObjectsState
  Set-ActionButtonState

}


function Get-FileButton($NameFilter, $InitialPath) {

  # Help function for the OpenFile Dialog window

  $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  if ($InitialPath -ne "") {
    $OpenFileDialog.InitialDirectory = $InitialPath
  }
  $OpenFileDialog.Filter = "$NameFilter"
  $OpenFileDialog.ShowHelp = $True
  $OpenFileDialog.ShowDialog() | Out-Null
  $OpenFileDialog.FileName

}


function Get-IPRange {
  
  # Create a list of IP's

  if ( (!([bool]($IPRangeFromText.Text -as [ipaddress]))) -or (!([bool]($IPRangeToText.Text -as [ipaddress]))) ) {
    # One of the values entered is not a correct IP
    [void][System.Windows.Forms.MessageBox]::Show("Unable to validate the IP.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
  }
  
  # IP operations
  $IP1 = ([System.Net.IPAddress]$($IPRangeFromText.Text)).GetAddressBytes()
  [Array]::Reverse($IP1)
  $IP1 = ([System.Net.IPAddress]($IP1 -join '.')).Address

  $IP2 = ([System.Net.IPAddress]$($IPRangeToText.Text)).GetAddressBytes()
  [Array]::Reverse($IP2)
  $IP2 = ([System.Net.IPAddress]($IP2 -join '.')).Address

  # Create the IP range
  $IPObjectList = @()
  for ($x = $IP1; $x -le $IP2; $x++) {
    $IP = ([System.Net.IPAddress]$x).GetAddressBytes()
    [Array]::Reverse($IP)
    $IPObjectList += ($IP -join '.')
  }

  if (@($IPObjectList).Count -eq 0) {
    [void][System.Windows.Forms.MessageBox]::Show("Unable to create a range with these values.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
  }

  if (@($IPObjectList).Count -gt 16384) {
    [void][System.Windows.Forms.MessageBox]::Show("Unable to create a range with more than 16384 values.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
  }

  $FormIPRange.Close()
  Add-ObjectListToGrid $IPObjectList ""  # Add the objects created to the grid

}


function Get-NewSequenceList {

  # Load a new Sequence List

  $SequencesListPath = Get-FileButton "Hydra Sequence List (*.lst)|*.lst|All files|*.*" $HydraSettingsPath

  if (($SequencesListPath -eq $Null) -or ($SequencesListPath -eq "")) { return }
  Set-ReloadSequenceList

}


function Get-ObjectsAD {

  # Get AD objects from a Query

  if (!(Get-Module ActiveDirectory)) {
    # Load the AD module if not already loaded
    Import-Module ActiveDirectory
  }
  
  # Run the Query defined in the "Query AD" window
  $ADObjList = Invoke-Expression -Command $ADQueryText.Text | where { $_ -like $ADQueryFilterText.Text }

  if ($ADObjList -eq $Null) {
    [System.Windows.Forms.MessageBox]::Show("Nothing found matching your query", "AD Query", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Stop) 
    return
  }

  $FormADQuery.Close()

  Add-ObjectListToGrid $ADObjList ""  # Add the objects found to the grid

}


function Get-ObjectsFile {

  # Read the content of a file and add every object to the grid

  $ObjFilePath = Get-FileButton "All files (*.*)|*.*" $LastDirObjects
  if (($ObjFilePath -eq $Null) -or ($ObjFilePath -eq "")) { return }

  $Separator = ";", ",", "|"  # Use Separators to split several objects defined on one line, and remove useless spaces
  [System.Collections.ArrayList]$ObjectList = @((Get-Content $ObjFilePath).Split($Separator).Trim())

  for ($i = $ObjectList.Count - 1; $i -ge 0; $i--) {
    # Remove empty objects to avoid empty rows
    if ($ObjectList[$i] -eq "") { $ObjectList.RemoveAt($i) }
  }  

  Add-ObjectListToGrid $ObjectList $ObjFilePath  # Add the objects found to the grid as well as their corresponding file
  $Script:LastDirObjects = Split-Path $ObjFilePath  # Save the last directory for registry user's settings on close

}


function Get-ObjectsManually {

  # Enter or paste a list of objects separated by separators

  $ObjectList = Read-InputBoxDialog "Objects" "Enter the list of Objects separated by a comma, semicolon or pipe:" ""
  if ($ObjectList -eq "") { return }

  $Separator = ";", ",", "|"  # Use Separators to split the objects and remove useless spaces
  [System.Collections.ArrayList]$ObjectList = @($ObjectList.Split($Separator).Trim())

  for ($i = $ObjectList.Count - 1; $i -ge 0; $i--) {
    # Remove empty objects to avoid empty rows
    if ($ObjectList[$i] -eq "") { $ObjectList.RemoveAt($i) }
  }  

  Add-ObjectListToGrid $ObjectList ""  # Add the objects to the grid

}


function Get-ObjectsPatse {

  # Paste the objects stored in the Clipboard to the grid

  $Clipboard = [System.Windows.Forms.Clipboard]::GetText()
  if ($Clipboard -eq "") { return }
  $Separator = ";", ",", "|", "`r", "`n", "`t"  # Use Separators to split the objects and remove useless spaces
  $Clipboard = @($Clipboard.Split($Separator).Trim())
  $ObjectList = @()

  foreach ($item in $Clipboard) {
    # Only add non-empty objects
    if ($item -ne "") { $ObjectList += $item }
  }  

  Add-ObjectListToGrid $ObjectList ""  # Add the objects to the grid

}


function Get-ObjectsSCCM($QueryTpye) {

  # Get Objects with a SCCM Query

  switch ($QueryTpye) {
    # Define the appropriate Query

    "Object" {
      $ObjectPattern = $SCCMQueryObjText.Text
      $WmiQuery = "
        Select DISTINCT *
        FROM SMS_R_System
        WHERE SMS_R_System.Name IS LIKE '$ObjectPattern%'"

      $WmiParams = @{
        'ComputerName' = $SCCM_ConfigMgrSiteServer
        'Namespace'    = "root\sms\site_$SCCM_SiteCode"
        'Query'        = $WmiQuery  
      }
    }

    "IP" {
      $IPPattern = $SCCMQueryIPText.Text
      $WmiQuery = "
        Select DISTINCT *
        FROM SMS_R_System
        WHERE SMS_R_System.IPAddresses IS LIKE '$IPPattern%'"

      $WmiParams = @{
        'ComputerName' = $SCCM_ConfigMgrSiteServer
        'Namespace'    = "root\sms\site_$SCCM_SiteCode"
        'Query'        = $WmiQuery  
      }
    }

    "Manual" {
      $WmiQuery = $SCCMQueryManualText.Text

      $WmiParams = @{
        'ComputerName' = $SCCM_ConfigMgrSiteServer
        'Namespace'    = "root\sms\site_$SCCM_SiteCode"
        'Query'        = $WmiQuery  
      }
    }

  }

  try {
    # Execute the Query
    $SCCMQueryResult = Get-WmiObject @WmiParams -ErrorAction Stop | select -ExpandProperty Name
  }

  catch {
    [System.Windows.Forms.MessageBox]::Show("Unable to connect to the SCCM server", "SCCM Query", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Stop)
    return
  }
  
  if ($SCCMQueryResult -eq "") {
    [System.Windows.Forms.MessageBox]::Show("Nothing found matching your query", "SCCM Query", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Stop) 
    return
  }

  $FormSCCMQuery.Close()

  Add-ObjectListToGrid $SCCMQueryResult ""  # Add the Objects found to the grid

}


function Get-RegistrySettings {

  # Read all the variables set in HKCU:\SOFTWARE\Hydra\5 and override the default ones

  $RegHydra = Get-ItemProperty HKCU:\SOFTWARE\Hydra\5 -ErrorAction SilentlyContinue | select * -ExcludeProperty PS* | ForEach-Object { $_.PSObject.Properties } | Select-Object Name, Value
  if ($RegHydra -ne $Null) {
    foreach ($RegEntry in $RegHydra) { 
      if ($RegEntry.Name -like ("Color_*")) {
        # Color variable found: set the HEX value
        $Colors.Set_Item(($RegEntry.Name -split ("_"))[1], $RegEntry.Value)
      }
      else {
        # Set the value found to the corresponding name
        Set-Variable -Name $RegEntry.Name -Value $RegEntry.Value -Scope Script -Force
      }
    }
  }

}


function Get-RowState($row, $TabIndex) {

  # Set the State of a row (values, colors,  depending on the Runspace result

  $xPID = $row.Cells[1].Value  # Runspace ID
  $SeqId = $row.Cells[8].Value  # Sequence ID

  try {
    if ($RunspaceCollection[$SeqId][$xPID].Runspace.IsCompleted) {
      # Check if the Runspace associated to the object has finished
      Get-RowState_ReturnedValue $row $xPID  # Get the Value returned by the Runspace
      if ($row.Cells[4].Value -eq -5) {
        # The row was in a Cancelling state: change it to Cancelled
        $row.Cells[4].Value = -6  # Set the State to -6: CANCELLED
        $row.Cells[2].Value = "Cancelled"
        $Row.Cells[0].Tag.PreviousStateComment = "Cancelled"
        $row.Cells[0].Tag.StepProtocol += @("   Cancelled")
        $row.Cells[3].Value = "CANCELLED"
        $row.DefaultCellStyle.BackColor = $row.Cells[5].Value
        if ($row.Cells[0].Tag.GroupID -eq "0") { $Row.Cells[0].ReadOnly = $False }
        $row.Cells[7].Value = $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked
        $Row.Cells[7].ReadOnly = $False   
        if ($row.Cells[7].Value -eq $False) { $Script:nbCheckedBoxes-- }
        $Script:ObjectsDone++
      }

      if ($row.Cells[4].Value -eq $Sequences[$SeqId].ScriptBlockCheckboxes.Count) {
        # The Runspace is completed and the Step ID equals the number of steps: the sequence has finished
        $row.Cells[4].Value = -1  # Set the State to -1: OK
        $row.DefaultCellStyle.BackColor = $row.Cells[5].Value
        if ($row.Cells[0].Tag.GroupID -eq "0") { $Row.Cells[0].ReadOnly = $False }  # Not member of a group: make the name of the object editable again
        $row.Cells[7].Value = $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked
        $Row.Cells[7].ReadOnly = $False
        if ($row.Cells[7].Value -eq $False) { $Script:nbCheckedBoxes-- }
        $Script:ObjectsDone++
      }

      if ($row.Cells[4].Value -lt 0) {
        # Close the protocol
        $row.Cells[0].Tag.StepProtocol += @("   Ended at $CurrentTime")
      }

      if ($Sequences[$SeqId].ScriptBlockPreLoad) {
        # The Step was a Pre-Load
        for ($i = 1; $i -lt $OutputDataGridSequence[$SeqId].Count; $i++) {
          # Loop in all objects in the sequence
          $OutputDataGridSequence[$SeqId][$i].Cells[0].Tag.SharedVariable = $row.Cells[0].Tag.SharedVariable  # Set the Shared Variable of the 1st objects to all other objects
          if ($($row.Cells[4].Value) -lt -1) {
            # The Pre-Load was not OK: Cancel all the steps of the SeqId objects
            $OutputDataGridSequence[$SeqId][$i].Cells[4].Value = -6
            $OutputDataGridSequence[$SeqId][$i].Cells[2].Value = "Cancelled - PreLoad not OK"
            $OutputDataGridSequence[$SeqId][$i].Cells[0].Tag.PreviousStateComment = "Cancelled"
            $OutputDataGridSequence[$SeqId][$i].Cells[0].Tag.StepProtocol += @("   Cancelled")
            $OutputDataGridSequence[$SeqId][$i].Cells[3].Value = "CANCELLED"
            $OutputDataGridSequence[$SeqId][$i].DefaultCellStyle.BackColor = $Colors.Get_Item("CANCELLED")
            if ($OutputDataGridSequence[$SeqId][$i].Cells[0].Tag.GroupID -eq "0") { $OutputDataGridSequence[$SeqId][$i].Cells[0].ReadOnly = $False }
            $OutputDataGridSequence[$SeqId][$i].Cells[7].Value = $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked
            $OutputDataGridSequence[$SeqId][$i].Cells[7].ReadOnly = $False   
            if ($OutputDataGridSequence[$SeqId][$i].Cells[7].Value -eq $False) { $Script:nbCheckedBoxes-- }
            $Script:ObjectsDone++
          }
        }

      }

    }

  }

  catch {
    # The state of the Runspace couldn't be queried: it can be due to a script error, an object returning a mismatch, or a bug in the Hydra code
    $row.Cells[4].Value = -6  # Set the row in the Cancelled state
    $row.Cells[2].Value = "Runspace Error - Cancelled"
    $row.Cells[0].Tag.PreviousStateComment = "Runspace Error - Cancelled"
    $row.Cells[0].Tag.StepProtocol += @("   Runspace Error - Cancelled")
    $row.Cells[3].Value = "CANCELLED"
    $row.Cells[5].Value = $Colors.Get_Item("STOP")
    $row.DefaultCellStyle.BackColor = $row.Cells[5].Value
    $row.Cells[0].Tag.StepProtocol += @("   Ended at $CurrentTime")
    if ($row.Cells[0].Tag.GroupID -eq "0") { $Row.Cells[0].ReadOnly = $False }
    $row.Cells[7].Value = $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked
    $Row.Cells[7].ReadOnly = $False
    if ($row.Cells[7].Value -eq $False) { $Script:nbCheckedBoxes-- }
    $Script:ObjectsDone++
  }

}


function Get-RowState_ReturnedValue($row, $xPID) {

  # A Runspace has finished: get its returned object. The returned objects is a collections of values: State, Comment [,Color] [,Shared Variable] 
  # Values of the StepID: 
  #    0: Sequence not started ; >0: Step currently running ; -1: OK ; -2: STOP ; -3: BREAK ; -4: ERROR ; -5: CANCELLING ; -6: CANCELLED

  $SeqId = $row.Cells[8].Value  # Get the Sequence ID of the row passed as parameter

  try {
    $xReceive = $RunspaceCollection[$SeqId][$xPID].PowerShell.EndInvoke($RunspaceCollection[$SeqId][$xPID].Runspace)  # Use the EndInvoke methode to get the state of the Runspace
    $Script:RunspaceCollection[$SeqId][$xPID].PowerShell.Dispose()  # Dispose the Runspace
    $JobResultState = $xReceive[0]  # The 1st mandatory returned object is the State ("OK", "STOP", "BREAK")
    $JobResultComment = $xReceive[1]  # The 2nd mandatory returned object is the Comment to print
  }
  catch {
    # Errors were found getting the object returned by the Runspace
    $JobResultState = "ERROR" 
    $JobResultComment = "Error in Task (Not enough objects returned: enable Debug for details)"
    if ($DebugMode -eq 5) { 
      Write-Host "`n $($row.Cells[0].Value): Not enough Objects returned" 
      Write-DebugReceiveOutput $xReceive
    }
    $xReceive = @("ERROR", "", $Colors.Get_Item("CANCELLED"))  # Recreate the $xReceive object to set the row in Error and paint it in the color of "Cancelled" 
    $row.Cells[4].Value = -4  # Set the value of the current step ID to "ERROR"
  }

  if ($xReceive.Count -gt 4) {
    # More than 4 values were found in the object returned by the Runspace
    $JobResultState = "ERROR" 
    $JobResultComment = "Error in Task (Too much objects returned: enable Debug for details)"
    $row.Cells[4].Value = -4  # Set the value of the current step ID to "ERROR"
    if ($DebugMode -eq 5) {
      Write-Host "`n $($row.Cells[0].Value): $($xReceive.Count) Objects returned (too much)"
      Write-DebugReceiveOutput $xReceive
    }
    $xReceive = @("ERROR", "", $Colors.Get_Item("CANCELLED"))  # Recreate the $xReceive object to set the row in Error and paint it in the color of "Cancelled"
  }
  elseif ($xReceive[0] -NotIn @("OK", "BREAK", "STOP", "ERROR")) {
    # The STATE keyword received is unknown
    $JobResultState = "ERROR" 
    $JobResultComment = "Error in Task (Wrong keyword returned: enable Debug for details)"
    $row.Cells[4].Value = -4  # Set the value of the current step ID to "ERROR"
    if ($DebugMode -eq 5) { 
      Write-Host "`n $($row.Cells[0].Value): Wrong keyword returned: $($xReceive[0])"
      Write-DebugReceiveOutput $xReceive
    }
    $xReceive = @("ERROR", "", $Colors.Get_Item("CANCELLED"))  # Recreate the $xReceive object to set the row in Error and paint it in the color of "Cancelled"    
  }

  if (($xReceive[0] -eq "STOP") -and ($row.Cells[4].Value -ge 0)) {
    # The State returned is STOP and the sequence was still running
    $JobResultState = "STOP at step $($row.Cells[4].Value)"  # Modify the State to print
    $row.Cells[4].Value = -2  # Set the value of the current step ID to "STOP"
  }
  
  if (($xReceive[0] -eq "BREAK") -and ($row.Cells[4].Value -ge 0)) {
    # The State returned is BREAK and the sequence was still running
    $JobResultState = "BREAK at step $($row.Cells[4].Value)"  # Modify the State to print
    $row.Cells[4].Value = -3  # Set the value of the current step ID to "BREAK"
  }

  if (($xReceive.Count -ge 3) -and ($xReceive[2] -ne $Null)) {
    # 3 or more variables returned, and a color has been defined
    $ColorsReturned = $xReceive[2].split("|")
    try {
      [windows.media.color]$($ColorsReturned[0]) | Out-Null  # Check if the 3rd variable is a HTML color name
      $IsHTMLColor = $True
    }
    catch {
      $IsHTMLColor = $False
    }
    if (($ColorsReturned[0] -match '#ff(([0-9a-f]{6}))\b') -or ($IsHTMLColor -eq $True)) {
      # The 3rd variable is a HTML color value (#FFxxxxxx) or name
      $BackgroundColor = $ColorsReturned[0]  # Set the background color of the cell to the 3rd value returned
    }
    else {
      $BackgroundColor = $Colors.Get_Item($xReceive[0])  # set the default color of the State returned (1st variable)
      if (($DebugMode -eq 5) -and ($ColorsReturned[0] -ne "")) {
        Write-Host "`nWrong color value: $($ColorsReturned[0]) is not a valid HTML color or in format #FFxxxxxx. Reset to default"
      }
    }
    if ($ColorsReturned.Count -gt 1) {
      # More colors arguments: get the 2nd one, the Cell Font Style
      $FontStyles = @("Regular", "Italic", "Bold", "Strikeout", "Underline") 
      if ($ColorsReturned[1] -in $FontStyles) {
        $row.Cells[2].Style.Font = New-Object Drawing.Font("Microsoft Sans Serif", 8.25, [Drawing.FontStyle]::$($ColorsReturned[1]))
      }
      else {
        $CellFontReturned = $ColorsReturned[1].split(',') | ForEach-Object { Invoke-Expression $_ }
        try {
          $row.Cells[2].Style.Font = New-Object Drawing.Font($CellFontReturned)
        }
        catch {
          $row.Cells[2].Style.Font = New-Object Drawing.Font("Microsoft Sans Serif", 8.25, [Drawing.FontStyle]::Regular)
          if ($DebugMode -eq 5) {
            Write-Host "`nWrong Font value: $($ColorsReturned[1]) is not a valid System.Drawing.Font. Reset to default"
          }
        }
      }
    }
    if ($ColorsReturned.Count -gt 2) {
      # More colors arguments: get the 3nd one, the Cell Font Color
      try {
        [system.windows.media.color]$($ColorsReturned[2]) | Out-Null  # Check if the 3rd variable is a HTML color name
        $IsHTMLColor = $True
      }
      catch {
        $IsHTMLColor = $False
      }
      if (($ColorsReturned[2] -match '#ff(([0-9a-f]{6}))\b') -or ($IsHTMLColor -eq $True)) {
        # The Cell Font color is a HTML color value (#FFxxxxxx) or name
        $row.Cells[2].Style.ForeColor = $ColorsReturned[2]  # Set the font color of the cell to the 3rd value returned
      }
    }
  }
  else { 
    $BackgroundColor = $Colors.Get_Item($xReceive[0])  # The 3rd variable is $Null: set the default color of the State returned (1st variable)
  }

  if ($xReceive.Count -eq 4) {
    # 4 variables returned (Shared variable)
    $row.Cells[0].Tag.SharedVariable = $xReceive[3]  # Set the Shared variable usable for the next steps
  }

  if ($row.Cells[4].Value -gt -5) {
    # Step ID > -5, the sequence is not cancelling/cancelled
    $row.Cells[2].Value = $JobResultComment  # Print the comment in the 2nd cell
    if ($JobResultComment -notlike "*Skipped*") {
      # The step hasn't been skipped
      $row.Cells[0].Tag.PreviousStateComment = $JobResultComment  # Set the comment on Tag of Cell[1] for further protocol use
    }
    $row.Cells[0].Tag.StepProtocol += @("   $JobResultComment")  # Extend the protocol
    $row.Cells[3].Value = $JobResultState  # Print the State
    $row.Cells[5].Value = $BackgroundColor  # Set the color to use for the row
  }
  if ($row.Cells[4].Value -lt 0) {
    # The sequence is finished
    $row.DefaultCellStyle.BackColor = $row.Cells[5].Value  # Paint the row with the defined color
    $row.Cells[7].Value = $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked  # Check or uncheck the checkbox depending on the user's settings
    $Row.Cells[7].ReadOnly = $False  # Re-enable the checkbox
    if ($row.Cells[7].Value -eq $False) { $Script:nbCheckedBoxes-- }
    if ($row.Cells[0].Tag.GroupID -eq "0") { $Row.Cells[0].ReadOnly = $False }  # Not member of a group: make the object's cell editable
    $Script:nbCheckedBoxes-- 
    $Script:ObjectsDone++  # Increase the objects proceeded
  }

  $Script:ConcurrentJobs[$SeqId]--  # Decrease the current jobs of the sequence ID
  if ($ConcurrentJobs[$SeqId] -eq 0) { $Script:ConcurrentJobs[$SeqId] = 1 }

}


function Get-Sequence($FileSeqPath, $SeqName) {

  # Parse and load a sequence.xml file

  # Reset the sequence panel objects and clear the panel
  $Script:SequencePanelCheckbox = @()
  $Script:SequencePanelLabel = @()
  $Script:SequencePanelVariable = @()
  $SequenceTasksPanel.Controls.Clear()

  $TimerInterval = $TimerIntervalDefault
  $Timer.Interval = $Script:TimerInterval

  if (!(Test-Path $FileSeqPath)) {
    # The file passed as argument is not existing
    Set-SequencePanelTitle "$SeqName`n`n`n" "Red"
    Set-SequencePanelLabel "  Missing: $FileSeqPath not found`n`n" "Italic" "Red" 0
    $Script:SequenceName = ""
    $ActionButton.Enabled = $False  # Disable the Start button
    return
  }

  try {
    # Load the file as an XML one
    $xmldata = New-Object System.Xml.XmlDocument
    $xmldata.Load((Get-ChildItem -Path $FileSeqPath).FullName)
  }
  catch [System.Xml.XmlException] {
    # XML errors detected
    Set-SequencePanelTitle "$SeqName`n`n`n" "Red"
    Set-SequencePanelLabel "  Error: XML parse error in $FileSeqPath`n`n" "Italic" "Red" 0
    $ActionButton.Enabled = $False  # Disable the Start button
    return 
  }

  $Script:SequencePath = ""
  $Script:SequenceAbsolutePath = ""
  $Script:SecurityCode = ""
  $FileSeqParentPath = Split-Path (Resolve-Path $FileSeqPath) -Parent  # Get the path of the sequence file
  $Err = $False

  $Script:ScriptBlockLoaded = New-Object 'System.Collections.Generic.List[System.Object]'  # Create a collection for the Step's ScriptBlocks
  $Script:ScriptBlockCommentLoaded = New-Object 'System.Collections.Generic.List[System.Object]'  # Create a collection for the Step's Comments
  # Reset some values to their default or empty values
  $MaxThreadsText.Text = $DefaultThreads 
  $MaxObj = 0 
  $Script:DisplayWarning = $False
  $Script:SendMail = $False
  New-Variable -Name MailServer -Value "" -Scope Script -Force
  New-Variable -Name MailFrom -Value "" -Scope Script -Force
  New-Variable -Name MailTo -Value "" -Scope Script -Force
  New-Variable -Name MailReplyTo -Value "" -Scope Script -Force

  $SeqPosition = 0  # Position of the Step for the task or PreLoad task

  # Get all the variables with a node name "parameter"
  $XMLSeqParam = ($xmldata.sequence | select parameter).parameter
  foreach ($XMLParam in $XMLSeqParam) {
    # Loop into the "parameters" found
    switch ($XMLParam.name) {
      # If the parameter found is known, set its value
      "sequencename" { $SeqName = $XMLParam.value }
      "warning" { if ($XMLParam.value -eq "yes") { $Script:DisplayWarning = $True } }
      "securitycode" { $Script:SecurityCode = $XMLParam.value }
      "maxthreads" { $MaxThreadsText.Text = $XMLParam.value }
      "maxobjects" { $MaxObj = $XMLParam.value }
      "sendmail" { if ($XMLParam.value -eq "yes") { $Script:SendMail = $True } }
      "timer" { if (([int]$XMLParam.value -ge 500) -and ([int]$XMLParam.value -le 30000)) { $Script:TimerInterval = [int]$XMLParam.value ; $Timer.Interval = $Script:TimerInterval } }
    }
    switch -wildcard ($XMLParam.name) {
      # A mail parameter has been detected
      "mail*" { New-Variable -Name $_ -Value $XMLParam.value -Scope Script -Force }
    }
  }

  # Get all the variables with a node name "importmodule"
  $XMLSeqMod = ($xmldata.sequence | select importmodule).importmodule
  $Script:SequenceImportModuleLoaded = @()  # Array for the modules
  foreach ($XMLModule in $XMLSeqMod) {
    # Loop into the "importmodule" found
    if (($XMLModule.type -eq "ImportPSSnapIn") -or ($XMLModule.type -eq "ImportPSModulesFromPath") -or ($XMLModule.type -eq "ImportPSModule")) {
      # The module type is known
      $Script:SequenceImportModuleLoaded += [PSCustomObject]@{  # Create a new module object with its type and name
        Type = $XMLModule.type
        Name = $XMLModule.name
      }
    }
  }

  # Get all the variables with a node name "variable"
  $XMLSeqVar = ($xmldata.sequence | select variable).variable
  $Script:SeqVariablesPos = @()  # Array for the variables
  foreach ($XMLVar in $XMLSeqVar) {
    # Loop into the "variable" found
    if ($XMLVar.type -in $VariableTypes) {
      # The variable Type is known
      $Script:SeqVariablesPos += [PSCustomObject]@{  # Create a new variable object with its type, name and value
        Type  = $XMLVar.type
        Name  = $XMLVar.name
        Value = $XMLVar.value 
      }
    }
  }

  # Get all the variables with a node name "preload"
  $XMLSeqPreload = ($xmldata.sequence | select preload).preload
  foreach ($XMLPreload in $XMLSeqPreload) {
    # Loop in all the preload tasks found
    $SeqPath = $XMLPreload.path  
    $SeqComment = $XMLPreload.comment
    $SeqFound = $False
    $SeqLocation = ""
    $TaskRelativeTo = [IO.Path]::Combine($FileSeqParentPath, $SeqPath)  # Built a path name based on the sequence.xml file and the PreLoad file path

    if (Test-Path $TaskRelativeTo) {
      # Search first the PreLoad file in the relative path of the sequence.xml file
      $SeqFound = $True 
      $SeqLocation = $TaskRelativeTo  # Set the Sequence location to this path
    }
    elseif (Test-Path $SeqPath) {
      # Search then in the path found in the node
      $SeqFound = $True 
      $SeqLocation = $SeqPath  # Set the Sequence location as defined in the node
    }

    if ($SeqFound -eq $True) {
      # The PreLoad is found
      $Error.Clear()
      # Load the PreLoad content as a ScriptBlock and assign it to the ScriptBlockLoaded array
      $Script:ScriptBlockLoaded += (Get-Command $SeqLocation -ErrorAction SilentlyContinue | select -ExpandProperty ScriptBlock -ErrorAction SilentlyContinue)
      if ($Error.Count -ne 0) {
        # A syntax error has been detected
        $ErrorMsg = (($Error[0].ToString() -split "`n")[0] -split ".ps1:")[1]  # Filter and print the error message
        Set-SequencePanelCheckbox "PreLoad $($SeqPosition+1)`n" $SeqPosition
        Set-SequencePanelLabel "  Error: error detected Line:$ErrorMsg`n" "Italic" "Red" $SeqPosition
        $Err = $True
      }
      else {
        # No syntax error detected
        $Script:ScriptBlockCommentLoaded += $SeqComment  # Assign the PreLoad comment to the ScriptBlockCommentLoaded array
        Set-SequencePanelCheckbox "PreLoad $($SeqPosition+1)`n" $SeqPosition  
        Set-SequencePanelLabel "  $SeqComment`n`n" "Regular" "Magenta" $SeqPosition
      }
    }
    else {
      # The PreLoad was not found
      Set-SequencePanelCheckbox "PreLoad $($SeqPosition+1)`n" $SeqPosition
      Set-SequencePanelLabel "  Missing: $SeqPath`n`n" "Italic" "Red" $SeqPosition
      $Err = $True
    }
    $SeqPosition++  # Increase the sequence Step position
  }

  # Get all the variables with a node name "task"
  $XMLSeqTask = ($xmldata.sequence | select task).task
  foreach ($XMLTask in $XMLSeqTask) {
    # Loop in all the tasks found
    $SeqPath = $XMLTask.path
    $SeqComment = $XMLTask.comment
    $SeqFound = $False
    $SeqLocation = ""
    $TaskRelativeTo = [IO.Path]::Combine($FileSeqParentPath, $SeqPath)  # Built a path name based on the sequence.xml file and the Task file path

    if (Test-Path $TaskRelativeTo) {
      # Search first the Task file in the relative path of the sequence.xml file
      $SeqFound = $True 
      $SeqLocation = $TaskRelativeTo  # Set the Sequence location to this path
    }
    elseif (Test-Path $SeqPath) {
      # Search then in the path found in the node
      $SeqFound = $True 
      $SeqLocation = $SeqPath  # Set the Sequence location as defined in the node
    }

    if ($SeqFound -eq $True) {
      # The Task is found
      $Error.Clear()
      # Load the Task content as a ScriptBlock and assign it to the ScriptBlockLoaded array
      $Script:ScriptBlockLoaded += (Get-Command $SeqLocation -ErrorAction SilentlyContinue | select -ExpandProperty ScriptBlock -ErrorAction SilentlyContinue)
      if ($Error.Count -ne 0) {
        # A syntax error has been detected
        $ErrorMsg = (($Error[0].ToString() -split "`n")[0] -split ".ps1:")[1]  # Filter and print the error message
        Set-SequencePanelCheckbox "Step $($SeqPosition+1)`n" $SeqPosition
        Set-SequencePanelLabel "  Error: error detected Line:$ErrorMsg`n" "Italic" "Red" $SeqPosition
        $Err = $True
      }
      else {
        # No syntax error detected
        $Script:ScriptBlockCommentLoaded += $SeqComment
        Set-SequencePanelCheckbox "Step $($SeqPosition+1)`n" $SeqPosition
        Set-SequencePanelLabel "  $SeqComment`n`n" "Regular" "Green" $SeqPosition
      }
    }
    else {
      # The Task was not found
      Set-SequencePanelCheckbox "Step $($SeqPosition+1)`n" $SeqPosition
      Set-SequencePanelLabel "  Missing: $SeqPath`n`n" "Italic" "Red" $SeqPosition
      $Err = $True
    }
    $SeqPosition++  # Increase the sequence Step position
  }

  Set-SequencePanelTitle "$SeqName`n`n`n" "Black"  # Set the Sequence name in the sequence panel

  $Script:SequenceLoaded = !($Err)  # No sequence loaded if error found
  $Script:SequenceName = $SeqName
  $Script:SequencePath = Split-Path $FileSeqPath -Parent
  $Script:SequenceAbsolutePath = $FileSeqParentPath
  $Script:SequenceFullPath = $FileSeqPath
  $Script:MaxSteps = $ScriptBlockLoaded.Count  # Set the number of Steps of the loaded sequence
  $Script:MaxObjects = $MaxObj

  Set-ActionButtonState

}


function Get-SequenceFileManual {

  # Loads a Sequence manually and creates an entry in the Sequence Tree

  # Select the file to process
  $SeqFilePath = Get-FileButton "Hydra Sequence (*.sequence.xml)|*.sequence.xml|All files|*.*" $LastDirSequences
  
  if (($SeqFilePath -eq $Null) -or ($SeqFilePath -eq "")) { return }
  $Script:LastDirSequences = Split-Path $SeqFilePath  # Set the variable $LastDirSequences to the folder of the sequence choosen. This will be reused as default folder for the next manual load

  $ManuallyLoadedSeq = $SequencesTreeView.Nodes | where { $_.Name -like "*Manually Loaded*" }  # Check if the Sequences Tree already has a parent node "Manually Loaded"

  if ($ManuallyLoadedSeq -eq $Null) {
    # The parent node "Manually Loaded" doesn't exist and is created
    $SequenceListRootNode = New-Object System.Windows.Forms.TreeNode
    $SequenceListRootNode.Text = "Manually Loaded"
    $SequenceListRootNode.Name = "Manually Loaded"
    [void]$SequencesTreeView.Nodes.Add($SequenceListRootNode)  # Add the new parent node "Manually Loaded" at the bottom of the Sequence tree
    $properties = @{'SeqName' = "----- Manually Loaded -----"; 'SeqPath' = "" }
    $object = New-Object –TypeName PSObject –Prop $properties
    $Script:SequenceList += $object  # The properties of this new parent node are added in the $SequenceList array
  }

  $SequenceListRootNode = $SequencesTreeView.Nodes | where { $_.Name -like "*Manually Loaded*" }  # Connects to "Manually Loaded"
  $SequenceListSubNode = New-Object -TypeName System.Windows.Forms.TreeNode
  $xmldata = New-Object System.Xml.XmlDocument  # Creates a new XML object
  $SeqName = $SeqFilePath
  try {
    # Search for a paramater "sequencename" in the .sequence.xml file selected
    $xmldata.Load((Get-ChildItem -Path $SeqFilePath).FullName)
    ($xmldata.sequence | select parameter).parameter | foreach { if ($_.Name -eq "sequencename") { $SeqName = " $($_.Value) ($SeqFilePath)" } }
  }
  catch [System.Xml.XmlException] {
    # if it's not found, the name of the Sequence will be the path of the file
  }
  $SequenceListSubNode.Text = $SeqName 
  $SequenceListSubNode.Tag = $($SequencesTreeView.Nodes.Nodes.Count + $SequencesTreeView.Nodes.Count)                   
  [void]$SequenceListRootNode.Nodes.Add($SequenceListSubNode)  # Add the new node $SequenceListSubNode in the "Manually Loaded" section
  $properties = @{'SeqName' = $SeqName; 'SeqPath' = $SeqFilePath }
  $object = New-Object –TypeName PSObject –Prop $properties
  $Script:SequenceList += $object  # The properties of this new node are added in the $SequenceList array
  
  $SequencesTreeView.SelectedNode = $SequenceListSubNode  # Select this new node
  
  if ($FormSettingsSequenceExpandedRadioButton.Checked) {
    # Depending on the user's settings, expand or collapse the Sequence Tree  
    $SequenceListRootNode.Expand()
  }
  else {
    $SequenceListRootNode.Collapse()
  }
  $Script:SelectionChanged = $True

}


function Import-Group {

  # Import Group(s) from a XML bundle

  # Get the file to open: a group or a grouplist
  $FileToImport = Get-FileButton "Hydra Groups Files(*.group.xml,*.grouplist)|*.group.xml;*.grouplist|Hydra Group (*.group.xml)|*.group.xml|Hydra Group List (*.grouplist)|*.grouplist|All files (*.*)|*.*" $LastDirImportGroup
  if (($FileToImport -eq "") -or ($FileToImport -eq $Null)) { return }
  $Script:LastDirImportGroup = Split-Path $FileToImport  # Set the variable $LastDirImportGroup to the folder of the sequence choosen. This will be reused as default folder for the next import

  $ErrorReturn = @()
  $NonImported = @()
  if ((Get-Content $FileToImport -First 1) -like "*Hydra Group List*") {
    # The file loaded is a grouplist
    foreach ($Path in (Get-Content $FileToImport)) {
      # Parse the file and import all files contained in it
      $Path = Join-Path $LastDirImportGroup $Path
      if ($Path -like "*Hydra Group List*") { continue }
      if (Test-Path $Path) { 
        $ErrorReturn += Import-GroupSingle $Path $True  # If an import fails, the $ErrorReturn won't be $Null
      }
      else {
        $NonImported += $Path
      }
    }
  }
  else {
    # A single group file has been chosen
    Import-GroupSingle $FileToImport $False
    $SequencesTreeView.SelectedNode = $SequencesTreeView.Nodes[0]  # Reset the sequence
    $Script:SelectionChanged = $True
    $Script:SequenceLoaded = $False
    return
  }

  if ($ErrorReturn -ne $Null) {
    # $ErrorReturn is not empty: some groups to import were already existing
    [void][System.Windows.Forms.MessageBox]::Show("The following Groups are already assigned and have been skipped:`r`n`r`n$($ErrorReturn -join "`r`n")", "Groups", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Stop) 
  }

  if ($NonImported -ne $Null) {
    # $NonImported is not empty: some groups to import were not found
    [void][System.Windows.Forms.MessageBox]::Show("The following Groups could not be found:`r`n`r`n$($NonImported -join "`r`n")", "Groups", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Stop) 
  }

  $SequencesTreeView.SelectedNode = $SequencesTreeView.Nodes[0]  # Reset the sequence
  $Script:SelectionChanged = $True
  $Script:SequenceLoaded = $False

}


function Import-GroupSingle($FileToImport, $Silent) {

  # Loads an XML file and set all attributes to the corresponding group

  Import-Clixml $FileToImport | foreach { Set-Variable $_.Name $_.Value }  # Reads all the variables stored in the XML file

  if ($GroupExported -in $GroupsUsed) {
    # Check if the Group name found in the file is already in use
    if (!($Silent)) {
      # If yes, and if allowed, a Warning is displayed
      [void][System.Windows.Forms.MessageBox]::Show("The Group Name '$GroupExported' is already assigned.", "Group Name", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Stop)
    }
    return $GroupExported  # Returns the name of the group to add in $ErrorReturn in the function Import-Group 
  }

  # The Group name found does not exist: increase all arrays needed for a new Sequence and add the content of the respective variables
  
  $NewSeq = New-Object -TypeName PSObject
  $NewSeq | Add-Member –MemberType NoteProperty –Name ScriptBlock –Value @($ScriptBlockExport)
  $NewSeq | Add-Member –MemberType NoteProperty –Name ScriptBlockComment –Value @($ScriptBlockCommentExport)
  $NewSeq | Add-Member –MemberType NoteProperty –Name ScriptBlockPreLoad –Value $False
  $NewSeq | Add-Member –MemberType NoteProperty –Name ScriptBlockVariable –Value @($ScriptBlockVariableExport)
  $NewSeq | Add-Member –MemberType NoteProperty –Name ScriptBlockModule –Value @($ScriptBlockModuleExport)
  $NewSeq | Add-Member –MemberType NoteProperty –Name ScriptBlockCheckboxes –Value @($ScriptBlockCheckboxesExport)
  $NewSeq | Add-Member –MemberType NoteProperty –Name SequenceScheduler –Value $SequenceSchedulerExport
  $NewSeq | Add-Member –MemberType NoteProperty –Name SequenceSendMail –Value $SequenceSendMailExport
  $NewSeq | Add-Member –MemberType NoteProperty –Name MaxCheckedObjects –Value $MaxCheckedObjectsExport
  $NewSeq | Add-Member –MemberType NoteProperty –Name SequenceLabel –Value $SequenceLabelExport
  $NewSeq | Add-Member –MemberType NoteProperty –Name MaxThreads –Value $MaxThreadsExport
  $NewSeq | Add-Member –MemberType NoteProperty –Name BelongsToGroup –Value $BelongsToGroupExport
  $NewSeq | Add-Member –MemberType NoteProperty –Name SequenceSchedulerExpired –Value $SequenceSchedulerExpiredExport
  $NewSeq | Add-Member –MemberType NoteProperty –Name SecurityCode –Value $SecurityCodeExport
  $NewSeq | Add-Member –MemberType NoteProperty –Name DisplayWarning –Value $DisplayWarningExport
  $Script:Sequences += $NewSeq

  $Script:RunspaceCollection += , @()
  $Script:RunspacePool += , @() 
  $Script:JobNb += , @()
  $Script:ConcurrentJobs += , @()

  $Script:SequenceTabIndex += $OutputDataGrid.Tag.TabPageIndex
  $Script:OutputDataGridSequence += , @()
  $Script:GroupsUsed += $GroupExported

  $DataGridViewCellStyleBold = New-Object System.Windows.Forms.DataGridViewCellStyle
  $DataGridViewCellStyleBold.Alignment = 16
  $DataGridViewCellStyleBold.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [Drawing.FontStyle]::"Bold", 3, 0)
  $DataGridViewCellStyleBold.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
  $DataGridViewCellStyleBold.SelectionBackColor = [System.Drawing.Color]::FromArgb(255, 51, 153, 255)
  $DataGridViewCellStyleBold.SelectionForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 255)

  $DataGridViewCellStyleRegular = New-Object System.Windows.Forms.DataGridViewCellStyle
  $DataGridViewCellStyleRegular.Alignment = 16
  $DataGridViewCellStyleRegular.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [Drawing.FontStyle]::"Regular", 3, 0)
  $DataGridViewCellStyleRegular.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
  $DataGridViewCellStyleRegular.SelectionBackColor = [System.Drawing.Color]::FromArgb(255, 51, 153, 255)
  $DataGridViewCellStyleRegular.SelectionForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 255)

  if (($Sequences[$Sequences.Count - 1].SequenceScheduler) -as [DateTime]) {
    # A scheduler has been set
    $TimeDiff = New-TimeSpan $(Get-Date) $Sequences[$Sequences.Count - 1].SequenceScheduler
    if ($TimeDiff.TotalSeconds -le 1) {
      # Timer expired
      $Script:Sequences[$Sequences.Count - 1].SequenceScheduler = 0
      $PendingText = "Pending"
    }  
    else {
      $PendingText = ($Sequences[$Sequences.Count - 1].SequenceScheduler).ToLongTimeString()
    }
  }
  else {
    # No timer defined
    $PendingText = "Pending"
  }

  $SeqId = $Sequences.Count - 1
  foreach ($Obj in $ObjectsList) {
    # Add the objects in the current grid, with pending values
    $RowID = $OutputDataGrid.Rows.Add($Obj, "0", "Sequence assigned: $SequenceLabelExport", $PendingText, 0, "", "", $True, $SeqId, "$GroupExported ($MaxThreadsExport)")
    $OutputDataGrid.Rows[$RowID].Cells[0].Style = $DataGridViewCellStyleBold
    $OutputDataGrid.Rows[$RowID].Cells[2].Style = $DataGridViewCellStyleRegular
    $OutputDataGrid.Rows[$RowID].Cells[0].ReadOnly = $True  # Objects names can't be modified in Groups
    $ObjectOptions = New-Object -TypeName PSObject
    $ObjectOptions | Add-Member –MemberType NoteProperty –Name GroupID –Value $GroupExported
    $ObjectOptions | Add-Member –MemberType NoteProperty –Name PreviousStateComment –Value ""
    $ObjectOptions | Add-Member –MemberType NoteProperty –Name StepProtocol –Value $Null
    $ObjectOptions | Add-Member –MemberType NoteProperty –Name SharedVariable –Value $Null
    $OutputDataGrid.Rows[$RowID].Cells[0].Tag = $ObjectOptions
    $OutputDataGrid.Rows[$RowID].Cells[7].ReadOnly = $False
    $OutputDataGrid.Rows[$RowID].DefaultCellStyle.BackColor = "White"
    $Script:OutputDataGridSequence[$SeqId] += $OutputDataGrid.Rows[$RowID]  # Add the row defined in the $OutputDataGridSequence of the current Sequence
  }

  $Script:SequenceLoaded = $True
  Get-CountCheckboxes

  return $Null

}


function Import-Tabs {

  # Import Tab(s) from a .tabs file

  # Get the file to open
  $FileToImport = Get-FileButton "Hydra Tabs (*.tabs)|*.tabs|All files (*.*)|*.*" $LastDirImportGroup
  if (($FileToImport -eq "") -or ($FileToImport -eq $Null)) { return }
  $Script:LastDirImportGroup = Split-Path $FileToImport  # Set the variable $LastDirImportGroup to the folder of the sequence choosen. This will be reused as default folder for the next import

  foreach ($Line in (Get-Content $FileToImport)) {
    # Parse the file
    if ($Line -like "*Hydra Tabs*") { continue }  # Skip the header
    $LineSplit = $Line -split ";"  # Read and set the different Tab attributes
    $TabName = $LineSplit[0]
    $TabColor1 = $LineSplit[1]
    $TabColor2 = $LineSplit[2]
    $Objects = $LineSplit[3..$($LineSplit.Count - 1)]
    Set-NewTab
    $LastTab = $DataGridTabControl.TabCount - 1
    $GridTabIndex = $DataGridTabControl.TabPages[$LastTab].Tag.TabPageIndex
    $DataGridTabControl.TabPages[$LastTab].Text = "  $TabName  "
    $DataGridTabControl.TabPages[$LastTab].Tag.ColorSelected = $TabColor1
    $DataGridTabControl.TabPages[$LastTab].Tag.ColorUnSelected = $TabColor2
    $OutputDataGrid = $OutputDataGridTab[$GridTabIndex]
    Add-ObjectListToGrid $Objects ""
  }

  $DataGridTabControl.SelectedIndex = $DataGridTabControl.TabCount - 1
  $OutputDataGrid.ClearSelection()

}


function Remove-Tab {

  # Remove a Tab
  
  $RunningTask = $OutputDataGrid.Rows | where { [int]($_.Cells[4].Value) -gt 0 }  # Check if some objects are in a runing state and exits if any
  if (@($RunningTask).Count -gt 0) {
    [System.Windows.Forms.MessageBox]::Show("Unable to remove a Tab while Sequences are running.`n`r", "Tab", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Stop)
    return
  }

  $TabToDelete = $DataGridTabControl.SelectedTab
  Clear-Grid
  $DataGridTabControl.TabPages.Remove($TabToDelete)

}


function Rename-Tab($TabIndex) {

  # Rename the Tab 

  $NewTabName = (Read-InputBoxDialog "Tab" "Set the new Tab Name:" "")
  if ($NewTabName -eq "") { return }
  $DataGridTabControl.TabPages[$TabIndex].Text = "  $NewTabName  "

}


function Reset-AllObjects {

  # Recreate the Objects list of the selected tab with the Objects name only

  $AllObjects = @()
  for ($i = 0; $i -lt $OutputDataGrid.RowCount - 1; $i++) { $AllObjects += $OutputDataGrid.Rows[$i].Cells[0].Value }  # Get the Objects Name only
  $AllObjectsFile = @()
  for ($i = 0; $i -lt $OutputDataGrid.RowCount - 1; $i++) { $AllObjectsFile += $OutputDataGrid.Rows[$i].Cells[6].Value }  # Get the Objects' corresponding files they belong
  Clear-Grid  # Clear the grid
  Add-ObjectListToGrid $AllObjects ""  # Recreate the list with the prior saved names
  for ($i = 0; $i -lt $OutputDataGrid.RowCount - 2; $i++) { $OutputDataGrid.Rows[$i].Cells[6].Value = $AllObjectsFile[$i] }  # Add the corresponding files names

}


function Reset-DefaultSettings {

  # Reset all settings to Default

  $ReallyReset = [System.Windows.Forms.MessageBox]::Show("Do you really want to reset all settings to the default values ?`r`nThis will close this session of Hydra.", "WARNING", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
  if ($ReallyReset -eq "yes") {
    Remove-Item -Path HKCU:\Software\Hydra\5 -Force  # Delete and recreate HKCU:\Software\Hydra\5
    New-Item -Path 'HKCU:\Software\Hydra' -Name 5 | Out-Null
    Set-DefaultSettings  # Set the Settings to default 
    $Script:ResetSettings = $True  # With this option, Hydra won't save anything in the Registry on exit
    $Form.Close()
  }

}


function Reset-Runspaces {

  # Parse the sequence previously ran and close and dispose their RunspacePool

  foreach ($Index in $SequencesToParse) {
    try {
      $Script:RunspacePool[$Index].Close()
      $Script:RunspacePool[$Index].Dispose()
      $Script:RunspaceCollection[$Index] = , @()
      $Script:RunspacePool[$Index] = , @()
    }
    catch {}
  }
  [System.GC]::Collect()  # Clear some memory parts using the Garbage Collection
  $Script:SequencesToParse = New-Object System.Collections.ArrayList  # Set the $SequencesToParse to an empty array to prepare the next sequences run

}


function Reset-SequenceArrays {

  # Empty Sequence arrays that aren't used anymore

  # Detect all Sequence ID's on all Tabs
  $SequenceIndex = for ($i = 0; $i -lt $DataGridTabControl.TabCount; $i++) { 
    $TabGridID = $DataGridTabControl.TabPages[$i].Tag.TabPageIndex
    $OutputDataGridTab[$TabGridID].Rows.Cells | where { $_.ColumnIndex -eq 8 } | select -ExpandProperty Value -Unique 
  }
  
  if ($SequenceIndex -eq $null) { return }  # No Sequence found
  
  $AllSequencesId = 1..$($Sequences.Count - 1)
  $SequenceArraysToDelete = Compare-Object -ReferenceObject $AllSequencesId -DifferenceObject $SequenceIndex -PassThru  # Match the difference between the Sequences ID's assigned and the ID's in the grids
  foreach ($item in @($SequenceArraysToDelete)) {
    # The difference is the Sequences not assigned to any object anymore: all corresponding arrays can be emptied
    try {
      $Script:Sequences[$item] = $Null
      $Script:OutputDataGridSequence[$item] = $Null
    }
    catch {}
  }

}


function Restore-CurrentSettings {

  # Reset the user's settings as they were if the user cancels the changes he made

  for ($i = 0; $i -le 4; $i++) {
    $FormSettingsPathsText[$i].Text = $CurrentSettings[$i]
  }
  $FormSettingsLogCheckBox.Checked = $CurrentSettings[5]
  for ($i = 0; $i -le 3; $i++) {
    $FormSettingsColorsButton[$i].BackColor = $CurrentSettings[$i + 6]
  }
  $FormSettingsColorsGUIBackLabel.BackColor = $CurrentSettings[10]
  $FormSettingsColorsGUIBackButton.BackColor = $CurrentSettings[11]
  $FormSettingsSplashScreenCheckBox.Checked = $CurrentSettings[12]
  $FormSettingsDebugScreenCheckBox.Checked = $CurrentSettings[13]
  $FormSettingsSequenceExpandedRadioButton.Checked = $CurrentSettings[14]
  $FormSettingsColorsGUIBackButton.BackColor = $CurrentSettings[15]
  $FormSettingsColorsGUISeqButton.BackColor = $CurrentSettings[16]
  $FormSettingsColorsGUISeqRunButton.BackColor = $CurrentSettings[17]
  $FormSettingsSequenceShowSearchRadioButton.Checked = $CurrentSettings[18]
  $FormSettingsSequenceShowHideRadioButton.Checked = $CurrentSettings[19]
  $FormSettingsSplashScreenCheckBox.Checked = $CurrentSettings[20]
  $FormSettingsDebugScreenCheckBox.Checked = $CurrentSettings[21]
  $FormSettingsSequenceExpandedRadioButton.Checked = $CurrentSettings[22]
  $FormSettingsSequenceCollapsedRadioButton.Checked = $CurrentSettings[23]
  $FormSettingsRowHeaderGroupBoxHiddenRadioButton.Checked = $CurrentSettings[24]
  $FormSettingsRowHeaderGroupBoxVisibleRadioButton.Checked = $CurrentSettings[25]
  $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked = $CurrentSettings[26]
  $FormSettingsCheckBoxesGroupBoxResetStateRadioButton.Checked = $CurrentSettings[27]
  for ($i = 0; $i -le 3; $i++) {
    $FormSettingsMailText[$i].Text = $CurrentSettings[$i + 28]
  }
  $FormSettingsGroupsWarningUncheckedRadioButton.Checked = $CurrentSettings[32]
  $FormSettingsGroupsWarningCheckedRadioButton.Checked = $CurrentSettings[33]
  $FormSettingsGroupsThreadsVisibleRadioButton.Checked = $CurrentSettings[34]
  $FormSettingsGroupsThreadsInvisibleRadioButton.Checked = $CurrentSettings[35]

}


function Save-CurrentSettings {

  # Save the state of the user's settings in case the user will cancel the process

  $Script:CurrentSettings = @($FormSettingsPathsText[0].Text, $FormSettingsPathsText[1].Text, $FormSettingsPathsText[2].Text, $FormSettingsPathsText[3].Text,
    $FormSettingsPathsText[4].Text, $FormSettingsLogCheckBox.Checked, $FormSettingsColorsButton[0].BackColor, $FormSettingsColorsButton[1].BackColor, 
    $FormSettingsColorsButton[2].BackColor, $FormSettingsColorsButton[3].BackColor, $FormSettingsColorsGUIBackLabel.BackColor, $FormSettingsColorsGUIBackButton.BackColor, 
    $FormSettingsSplashScreenCheckBox.Checked, $FormSettingsDebugScreenCheckBox.Checked, $FormSettingsSequenceExpandedRadioButton.Checked,
    $FormSettingsColorsGUIBackButton.BackColor, $FormSettingsColorsGUISeqButton.BackColor, $FormSettingsColorsGUISeqRunButton.BackColor, 
    $FormSettingsSequenceShowSearchRadioButton.Checked, $FormSettingsSequenceShowHideRadioButton.Checked, $FormSettingsSplashScreenCheckBox.Checked,
    $FormSettingsDebugScreenCheckBox.Checked, $FormSettingsSequenceExpandedRadioButton.Checked, $FormSettingsSequenceCollapsedRadioButton.Checked,
    $FormSettingsRowHeaderGroupBoxHiddenRadioButton.Checked, $FormSettingsRowHeaderGroupBoxVisibleRadioButton.Checked, $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked,
    $FormSettingsCheckBoxesGroupBoxResetStateRadioButton.Checked, $FormSettingsMailText[0].Text, $FormSettingsMailText[1].Text, $FormSettingsMailText[2].Text,
    $FormSettingsMailText[3].Text, $FormSettingsGroupsWarningUncheckedRadioButton.Checked, $FormSettingsGroupsWarningCheckedRadioButton.Checked,
    $FormSettingsGroupsThreadsVisibleRadioButton.Checked, $FormSettingsGroupsThreadsInvisibleRadioButton.Checked )

}


function Send-Email {

  # Send the state of the grid per email

  if (($EMailSMTPServer -eq "") -or ($EMailSendFrom -eq "") -or ($EMailSendTo -eq "") -or ($EMailReplyTo -eq "")) {
    # If parameters are missing, exit
    [void][System.Windows.Forms.MessageBox]::Show("Unable to find the e-mail parameters.`r`n`r`nEnter the parameters in the Settings panel.`r`n" , "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
  }

  # Create a HMTL content based on the options checked by the user, and add the header
  $HTMLExport = Export-CreateHTML $FormExportColCheckBox[1].Checked $FormExportColCheckBox[2].Checked $FormExportColCheckBox[3].Checked $FormExportColCheckBox[4].Checked $FormExportColorCheckBox.Checked $FormExportSelectionCheckBox.Checked $True
  $ToSend = ConvertTo-Html -Head $($HTMLExport[1]) -Body "<H2>Hydra Deployment Results</H2> $($HTMLExport[0])" | Out-String

  # Send the mail
  fSend-Mail $EMailSMTPServer $EMailSendFrom $EMailSendTo $EMailReplyTo "Hydra Deployment Results" $ToSend $True
  #Send-MailMessage -To $EMailSendTo -From $EMailSendFrom -Subject "Hydra Deployment Results" -Body $ToSend -BodyAsHTML -SmtpServer $EMailSMTPServer

}


function Send-MailLog {

  # Parse the finished Sequences and send an email if it was set in the sequence.xml file

  $MailToSend = $False

  foreach ($Item in $SequencesToParse) {
    # Parse all Sequences just finished to run
    if ($Sequences[$Item].SequenceSendMail) {
      # The Sequence $Item should send an email
      if ($MailToSend -eq $False) {
        $MailToSend = $True
        $OutputDataGrid.ClearSelection()  # Clear the selection on the first occurence of $MailToSend -eq $False
      }
      foreach ($row in $OutputDataGridSequence[$Item]) {
        # Add all the objects of the Sequence $Item in the Grid Selection
        $row.Selected = $True
      }
    }
  }

  if ($MailToSend) {
    # A Mail has to be sent with all selected objects' details
    $HTMLExport = Export-CreateHTML $True $True $True $True $True $True $True
    $ToSend = ConvertTo-Html -Head $($HTMLExport[1]) -Body "<H2>Hydra Deployment Results</H2> $($HTMLExport[0])" | Out-String
    fSend-Mail $mailserver $mailfrom $mailto $mailreplyto "Hydra Deployment Results" $ToSend $True
    #Send-MailMessage -To $mailto -From $mailfrom -Subject "Hydra Deployment Results" -Body $ToSend -BodyAsHTML -SmtpServer $mailserver
    $OutputDataGrid.ClearSelection()
  }

}


function Set-ActionButtonState {

  # Enable or disable the "Start" button depending on different criterias

  $GroupFound = $False
  foreach ($row in $OutputDataGrid.Rows) {
    if ($Row.Index -eq $OutputDataGrid.RowCount - 1) { continue }
    if ($row.Cells[0].Tag.GroupID -ne "0") { $GroupFound = $True ; break }
  }
  $OutputDataGrid.Columns[9].Visible = $GroupFound  # If Groups are defined, show the Group Column and the menu items related
  $ExportGroupMenu.Enabled = $GroupFound
  $SelectAllGroupMenu.Enabled = $GroupFound
  $DeSelectAllGroupMenu.Enabled = $GroupFound

  $CheckboxesFound = ($nbCheckedBoxes -gt 0)

  if ($CheckboxesFound -eq $False) {
    # If no object selected, disable the button
    $ActionButton.Enabled = $False
    return
  }

  $SequenceAssigned = $False
  if (!($SequenceLoaded)) {
    foreach ($row in $OutputDataGrid.Rows) {
      if (($row.Cells[7].Value) -and ($row.Cells[0].Tag.GroupID -ne "0")) { $SequenceAssigned = $True ; break }  # Groups found and objects checked
    }
  }

  if (($SequenceLoaded -eq $False) -and ($SequenceAssigned -eq $False)) {
    # Neither a sequence is loaded nor Groups with objects assigned
    $ActionButton.Enabled = $False
    return
  }

  $ReadyToInstall = $False
  foreach ($row in $OutputDataGrid.Rows) {
    if (($row.Cells[7].Value) -and ($row.Cells[4].Value -le 0)) { $ReadyToInstall = $True ; break }  # Objects selected not running
  }

  $ActionButton.Enabled = $ReadyToInstall

}


function Set-AssignSequenceToFreeObjects {

  # Parse the grid for objects ready to start the selected Sequence

  if ($SequencesTreeView.SelectedNode.Parent -eq $NULL) { return }  # No Sequence selected

  # Determine the objects checked, not belonging to a group, without sequence assigned or a sequence already finished
  $SelectedRows = @($OutputDataGridTab[$GridIndex].Rows | where { ($_.Cells[7].Value) -and ($_.Cells[0].Tag.GroupID -eq "0") -and (($_.Cells[8].Value -eq 0) -or ($_.Cells[4].Value -lt 0) ) }) 
  $VariablesReloaded = $False

  if ($SelectedRows.Count -eq 0) { return }  # No objects matching the query

  if ($SecurityCode -ne "") {
    $CodePrompt = Read-SecurityCode $SecurityCode $SequenceName
    if ($CodePrompt -ne "OK") {
      return "err"
    }
  }
  elseif ($DisplayWarning -eq $True) {
    # The sequence has the Warning option
    $ReallyDeploy = Read-StartSequence $SequenceName
    if ($ReallyDeploy -ne "OK") {
      return "err"
    }
  }

  if ($SequencePanelVariable.Count -eq 0) {
    # The current loaded sequence doesn't have any variable set: if there are some, they are kept
    $Script:VariablesQuery = Set-CurrentSequenceToObjectsVariables  # Query and set the variables, if any
  }
  else {
    # Variables already set: ask if they should be reused or reloaded
    $ReloadVariables = [System.Windows.Forms.MessageBox]::Show("Variables have been already defined for $SequenceName`r`nDo you want to reuse them ?", "WARNING", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($ReloadVariables -eq "no") {
      for ($j = 0; $j -lt @($SequencePanelVariable).Count; $j++) { $Script:SequencePanelVariable[$j].Text = "" }# Clear the variables set in the Sequence Panel
      $Script:SequencePanelVariable = @()
      $Script:VariablesQuery = Set-CurrentSequenceToObjectsVariables  # Query and set the variables
      $VariablesReloaded = $True
    }
  }

  if ($VariablesQuery -eq "error") {
    # The variable query has been cancelled, the sequence start is stopped
    return "err"
  }

  foreach ($row in $SelectedRows) {
    # Set the status of the objects
    $row.Cells[2].Value = "Sequence loaded: $SequenceName"
    $row.DefaultCellStyle.BackColor = "White"
    $row.Cells[0].Tag.GroupID = "0"  # Set the GroupId to 0: doesn't belong to any group
  }

  $Script:SchedulerTemp = 0
  $GroupInuse = (@($GroupsUsed).Count -gt 0) -and ($GroupsUsed -ne "0")  # Are some groups in use

  # Assign the current sequence to the free and checked objects. If groups are in use, or the sequence has changed, assign a new ID too (2nd parameter)
  Set-CurrentSequenceToObjects $SelectedRows ($SelectionChanged -or $GroupInuse -or $VariablesReloaded) $False

}


function Set-AssignSequenceToObjects($UseScheduler, $GroupToAssign = $False) {

  # Assign the current sequence to objects in a group

  if ($SequencePanelVariable.Count -eq 0) {
    # The current loaded sequence doesn't have any variable set: if there are some, they are kept
    $Script:VariablesQuery = Set-CurrentSequenceToObjectsVariables  # Query and set the variables, if any
  }

  if ($VariablesQuery -eq "error") {
    # The variable query has been cancelled, the sequence start is stopped
    return 
  }
  
  $NewGroup = $False

  if ($GroupToAssign -ne $False) {
    # A group has been set via a the Sequence Tree right click
    $OutputDataGrid.ClearSelection()
    for ($i = 0; $i -lt $OutputDataGrid.RowCount - 1; $i++) {
      if ($OutputDataGrid.Rows[$i].Cells[0].Tag.GroupID -eq $GroupToAssign) {
        # Select the objects matching the group
        $OutputDataGrid.Rows[$i].Cells[0].Selected = $True
      }
    }
  }
  else {
    # Define a new group
    $GroupToSet = (Read-InputBoxDialog "Group" "Set the Group Name to assign '$($SequenceName)':" "")
    if ($GroupToSet -eq "") { return }
    if ($GroupToSet -in $GroupsUsed) {
      # Check if the name is already given to another group
      [System.Windows.Forms.MessageBox]::Show("The Group Name '$GroupToSet' is already assigned.", "Group Name", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Stop)
      return
    }
    $Script:NewGroupName = $GroupToSet
    $Script:GroupsUsed += $NewGroupName  # Extend $GroupsUsed with the newly created group name

    $NewGroup = $True
  }

  $SelectedRows = @($OutputDataGrid.Rows | where { $_.Cells[0].Selected })  # Determine the objects selected

  if ($UseScheduler -eq $True) {
    # Ask for a Scheduler if needed
    $Scheduler = Read-DateTimePicker "Enter the start for $SequenceName"
    if ($Scheduler -eq "") { return }
    $Script:SchedulerTemp = $Scheduler
  }
  else {
    $Script:SchedulerTemp = 0  # No Scheduler set
  }

  $DataGridViewCellStyleBold = New-Object System.Windows.Forms.DataGridViewCellStyle
  $DataGridViewCellStyleBold.Alignment = 16
  $DataGridViewCellStyleBold.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [Drawing.FontStyle]::"Bold", 3, 0)
  $DataGridViewCellStyleBold.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
  $DataGridViewCellStyleBold.SelectionBackColor = [System.Drawing.Color]::FromArgb(255, 51, 153, 255)
  $DataGridViewCellStyleBold.SelectionForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 255)

  $DataGridViewCellStyleRegular = New-Object System.Windows.Forms.DataGridViewCellStyle
  $DataGridViewCellStyleRegular.Alignment = 16
  $DataGridViewCellStyleRegular.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [Drawing.FontStyle]::"Regular", 3, 0)
  $DataGridViewCellStyleRegular.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
  $DataGridViewCellStyleRegular.SelectionBackColor = [System.Drawing.Color]::FromArgb(255, 51, 153, 255)
  $DataGridViewCellStyleRegular.SelectionForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 255)

  foreach ($row in $SelectedRows) {
    # Assign the current sequence to the selected objects
    $row.Cells[2].Value = "Sequence assigned: $SequenceName"
    $row.DefaultCellStyle.BackColor = "White"
    $row.Cells[0].Style = $DataGridViewCellStyleBold
    $row.Cells[2].Style = $DataGridViewCellStyleRegular
    $row.Cells[0].ReadOnly = $True
    if (($DisplayWarning -eq $True) -or ($SecurityCode -ne "")) {
      $row.Cells[7].Value = ($GrpCheckedOnWarning -eq "True")  # Auto uncheck the objects on warning if the option is set in the user's settings
    }
    if ($NewGroup) {
      # A new group has been define
      $row.Cells[0].Tag.GroupID = $NewGroupName  # Set the name of the group on the Cell's Tag
      $row.Cells[9].Value = "$NewGroupName"  # Set the group name in the Group Column
      if ($GrpShowThreads -eq "True") { $row.Cells[9].Value += " ($($MaxThreadsText.Text))" }  # Add also the number of Threads if the option is set in the user's settings
    }
    $Script:SequenceAssigned = $True
    if ($SchedulerTemp -ne 0) {
      # If a scheduler is set, set the time in the Step Column
      $row.Cells[3].Value = $SchedulerTemp.ToLongTimeString() 
    }
    else {
      $row.Cells[3].Value = "Pending"
    }  
  }

  # Assign the current sequence to the checked objects. For the groups, assign a new ID (2nd parameter)
  Set-CurrentSequenceToObjects $SelectedRows $True $True

  if ($UseAScheduler) {
    # A scheduler has been set in the sequence.xml: set the time in the Step Column
    foreach ($row in $SelectedRows) { $row.Cells[3].Value = $SchedulerTemp.ToLongTimeString() }
  }

  if ($SequencePanelVariable.Count -ne 0) {
    # Clear the variables set in the Sequence Panel
    for ($j = 0; $j -lt @($SequencePanelVariable).Count; $j++) { $Script:SequencePanelVariable[$j].Text = "" }
    $Script:SequencePanelVariable = @()
  }

  Set-ActionButtonState

}


function Set-CellValue($IndexOfGrid, $Row, $Id1, $Id2, $Id3, $Id4, $Id5, $Id6, $Id8) {

  # 0:Objects ; 1: JobID ; 2: Results ; 3: Step ; 4: StepID ; 5: Color ; 6: FileSource ; 8: SequenceId

  if ($Id1 -ne "#") { $OutputDataGridTab[$IndexOfGrid].Rows[$Row].Cells[1].Value = $Id1 }
  if ($Id2 -ne "#") { $OutputDataGridTab[$IndexOfGrid].Rows[$Row].Cells[2].Value = $Id2 }
  if ($Id3 -ne "#") { $OutputDataGridTab[$IndexOfGrid].Rows[$Row].Cells[3].Value = $Id3 }
  if ($Id4 -ne "#") { $OutputDataGridTab[$IndexOfGrid].Rows[$Row].Cells[4].Value = $Id4 }
  if ($Id5 -ne "#") { $OutputDataGridTab[$IndexOfGrid].Rows[$Row].Cells[5].Value = $Id5 }
  if ($Id6 -ne "#") { $OutputDataGridTab[$IndexOfGrid].Rows[$Row].Cells[6].Value = $Id6 }
  if ($Id8 -ne "#") { $OutputDataGridTab[$IndexOfGrid].Rows[$Row].Cells[8].Value = $Id8 }

}


function Set-CheckAll ($Check) {

  # Check or uncheck all the object of the visible grid

  for ($i = 0; $i -lt $OutputDataGrid.RowCount - 1; $i++) {
    $OutputDataGrid.Rows[$i].Cells[7].Value = $Check
  }

  $OutputDataGrid.RefreshEdit()

  Get-CountCheckboxes

}


function Set-CloseForm {

  # Save all current settings (windows size and position, user's settings, colors,...) into the registry

  if ($ResetSettings -eq $True) { return }  # $ResetSettings set to $True: nothing will be saved to the registry 

  $PosRegistry = $True
  $WindowPositions = @($Form.Top, $Form.Left, $Form.Size.Width, $Form.Size.Height, $SplitContainer1.Size.Width, $SplitContainer1.Size.Height, 
    $SplitContainer1.SplitterDistance, $SplitContainer2.Size.Width, $SplitContainer2.Size.Height, $SplitContainer2.SplitterDistance)

  foreach ($Pos in $WindowPositions) {
    if (($Pos -lt -20) -or ($Pos -gt 3000)) {
      # The window position seems to be wrong, it won't be saved
      $PosRegistry = $False
    }
  }

  if ($PosRegistry) {
    # No window position issue: save all variable values
    if ($ShowSearchBox -eq "True") { $SequencesTreeViewTopPosition = 20 } else { $SequencesTreeViewTopPosition = 0 }
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosFormY" -Value $Form.Top -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosFormX" -Value $Form.Left -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosFormW" -Value $($Form.Width - 2 * $FormBorderSize) -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosFormH" -Value $($Form.Height - 2 * $FormBorderSize - $FormHeaderSize) -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosSplit1W" -Value $($SplitContainer1.Size.Width - 2 * $FormBorderSize - 15) -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosSplit1H" -Value $SplitContainer1.Size.Height -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosSplit1D" -Value $SplitContainer1.SplitterDistance -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosSplit2W" -Value $($SplitContainer2.Size.Width - 15) -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosSplit2H" -Value $SplitContainer2.Size.Height -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PosSplit2D" -Value $SplitContainer2.SplitterDistance -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqTreeTop" -Value $($SequencesTreeView.Top - $SequencesTreeViewTopPosition) -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqTreeLeft" -Value $SequencesTreeView.Left -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqTreeWidth" -Value $SequencesTreeView.Width -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqTreeHeight" -Value $($SequencesTreeView.Height + $SequencesTreeViewTopPosition) -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqPanelTop" -Value $SequenceTasksPanel.Top -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqPanelLeft" -Value $SequenceTasksPanel.Left -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqPanelWidth" -Value $SequenceTasksPanel.Width -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SeqPanelHeight" -Value $SequenceTasksPanel.Height -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PanelBottomTop" -Value $PanelBottom.Top -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "PanelBottomWidth" -Value $($PanelBottom.Width - 2 * $FormBorderSize - 15) -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "DataGridTabControlWidth" -Value $($DataGridTabControl.Width - 2 * $FormBorderSize - 15) -Type String -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "DataGridTabControlHeight" -Value $DataGridTabControl.Height -Type String -Force
  }

  Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "LastDirObjects" -Value $LastDirObjects -Force
  Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "LastDirSequences" -Value $LastDirSequences -Force
  Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "LastDirExportGroup" -Value $LastDirExportGroup -Force
  Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "LastDirImportGroup" -Value $LastDirImportGroup -Force
  Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "WelcomeScreen" -Value "False" -Force

}


function Set-CopyObjectsToTab($Move, $Tab) {

  # Copy or move the selected objects to another or new tab

  if ($Tab -eq -1) {
    # A new tab has to be created
    Set-NewTab
    $TabToCopy = $TabPageIndex  # Target Tab is the new created one
  }
  else {
    $TabToCopy = $Tab  # Target Tab is the one passed as argument
  }

  $SelectedObjects = $OutputDataGrid.SelectedCells | select -ExpandProperty Value
  $CurrentTab = $DataGridTabControl.SelectedTab.TabIndex  # Save the position of the current Tab
  $OutputDataGrid = $OutputDataGridTab[$TabToCopy]  # Move to the target Tab
  Add-ObjectListToGrid $SelectedObjects ""  # Add the objects and go back to the original Tab
  $OutputDataGrid = $OutputDataGridTab[$CurrentTab]

  if ($Move -eq $True) { Set-RightClick_SetNewSelectionFromGrid $False }  # If Move option, delete the objects

}


function Set-CurrentSequenceToObjects($ObjectsSelected, $NewID, $BelongsToGroup) {

  # Assign the sequence to the selected objects 

  if ($ObjectsSelected.Count -eq 0) { return }

  if ($NewID) {
    # A new Sequence ID has to be used: increase all arrays and set the respective attributes
    if ($UseAScheduler) { $Script:SchedulerTemp = $SchedulerVar } 

    $NewSeq = New-Object -TypeName PSObject
    $NewSeq | Add-Member –MemberType NoteProperty –Name ScriptBlock –Value @($ScriptBlockLoaded)
    $NewSeq | Add-Member –MemberType NoteProperty –Name ScriptBlockComment –Value @($ScriptBlockCommentLoaded)
    $NewSeq | Add-Member –MemberType NoteProperty –Name ScriptBlockPreLoad –Value $False
    $NewSeq | Add-Member –MemberType NoteProperty –Name ScriptBlockVariable –Value @($VariablesQuery)
    $NewSeq | Add-Member –MemberType NoteProperty –Name ScriptBlockModule –Value @($SequenceImportModuleLoaded)
    $NewSeq | Add-Member –MemberType NoteProperty –Name ScriptBlockCheckboxes –Value @()
    $NewSeq | Add-Member –MemberType NoteProperty –Name SequenceScheduler –Value $SchedulerTemp
    $NewSeq | Add-Member –MemberType NoteProperty –Name SequenceSchedulerExpired –Value $False
    $NewSeq | Add-Member –MemberType NoteProperty –Name SequenceSendMail –Value $SendMail
    $NewSeq | Add-Member –MemberType NoteProperty –Name MaxCheckedObjects –Value $MaxObjects
    $NewSeq | Add-Member –MemberType NoteProperty –Name SequenceLabel –Value $SequenceName
    $NewSeq | Add-Member –MemberType NoteProperty –Name MaxThreads –Value $MaxThreadsText.Text
    $NewSeq | Add-Member –MemberType NoteProperty –Name SecurityCode -Value $SecurityCode
    $NewSeq | Add-Member –MemberType NoteProperty –Name BelongsToGroup -Value $BelongsToGroup
    if ($DisplayWarning) {
      $NewSeq | Add-Member –MemberType NoteProperty –Name DisplayWarning -Value 1
    }
    else {
      $NewSeq | Add-Member –MemberType NoteProperty –Name DisplayWarning -Value 0
    }
    $Script:Sequences += $NewSeq

    $Script:OutputDataGridSequence += , @()
    $Script:RunspaceCollection += , @()
    $Script:RunspacePool += , @() 
    $Script:JobNb += , @()
    $Script:ConcurrentJobs += , @()
    $Script:SequenceTabIndex += $OutputDataGrid.Tag.TabPageIndex
    $Script:OutputDataGridSequence[$Sequences.Count - 1] = @()

  }

  $RowsInUse = @($OutputDataGridSequence[$Sequences.Count - 1] | select -ExpandProperty Index)

  foreach ($Row in $ObjectsSelected) {
    # Set or reset the cells values
    $Row.Cells[4].Value = -1  # Set the Sequence Step to -1, Pending
    $Row.Cells[8].Value = $Sequences.Count - 1  # Assign the last Sequence ID to the object
    if ($Row.Index -notin $RowsInUse) {
      # A new object has been added to the Sequence
      $Script:OutputDataGridSequence[$Sequences.Count - 1] += $Row  # Add the row to the $OutputDataGridSequence associated to the Sequence ID
    }
  }

  # Recreate the ScriptBlockCheckboxes array to avoid that a pointer is created (Reference Type): solves the issue with multiple start of a Sequence with Preload
  $Script:Sequences[$Sequences.Count - 1].ScriptBlockCheckboxes = @()
  for ($i = 0; $i -lt $SequencePanelCheckbox.Count; $i++) {
    $Script:Sequences[$Sequences.Count - 1].ScriptBlockCheckboxes += New-Object System.Windows.Forms.CheckBox
    $Script:Sequences[$Sequences.Count - 1].ScriptBlockCheckboxes[$i].ForeColor = $SequencePanelCheckbox[$i].ForeColor 
    $Script:Sequences[$Sequences.Count - 1].ScriptBlockCheckboxes[$i].Location = New-Object System.Drawing.Size($SequencePanelCheckbox[$i].Left, $SequencePanelCheckbox[$i].Top)
    $Script:Sequences[$Sequences.Count - 1].ScriptBlockCheckboxes[$i].Checked = $SequencePanelCheckbox[$i].Checked
    $Script:Sequences[$Sequences.Count - 1].ScriptBlockCheckboxes[$i].AutoSize = $True
    $Script:Sequences[$Sequences.Count - 1].ScriptBlockCheckboxes[$i].MaximumSize = New-Object System.Drawing.Size(500, 15)
    $Script:Sequences[$Sequences.Count - 1].ScriptBlockCheckboxes[$i].Font = $SequencePanelCheckbox[$i].Font 
    $Script:Sequences[$Sequences.Count - 1].ScriptBlockCheckboxes[$i].Text = $SequencePanelCheckbox[$i].Text
  }

}


function Set-CurrentSequenceToObjectsVariables {

  # Query the value of each variable define in the sequence.xml

  if ($SequencePanelVariable.Count -ne 0) { return }  # No variable set: exit

  $VarP = 0  # Counter for the Y position of the variables values in the sequence panel
  $Script:UseAScheduler = $False 

  for ($i = 1; $i -le $nbVariableTypes; $i++) {
    # Define a hash for each variable type: it will store the name of the variables and their values in pairs
    $SeqVariableHash[$i] = @{}
  }

  for ($i = 0; $i -lt $SeqVariablesPos.Count; $i++) {
    # Loop through the variables defined in the sequence.xml
    $TypePos = $VariableTypes.IndexOf($SeqVariablesPos[$i].Type.ToLower())  # Determine the index in $VariableTypes of the variable type: the position is set in the function Set-StartupSettings_SeqVariables
    $SeqVariables = $SeqVariablesPos[$i].Value
    $VariableQuery = Invoke-Expression -Command $VariableCommand[$TypePos]  # Execute the command associated to the index above and store the result in $VariableQuery
    if ($VariableQuery -eq "") {
      # Query cancelled
      [System.Windows.Forms.MessageBox]::Show("  Process cancelled  ", $VariableTypes[$TypePos], [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Stop)
      for ($j = 0; $j -lt @($SequencePanelVariable).Count; $j++) { $Script:SequencePanelVariable[$j].Text = "" }  # Clean up the variables part in the sequence panel
      $Script:SequencePanelVariable = @()
      return "error"
    }
    $VariableName = $SeqVariablesPos[$i].Name  # Get the name of the variable defined in the sequence.xml and print it with the value returned above
    if ($VariableTypes[$TypePos] -ne "secretinputbox") {
      # If it is an secretinputbox, only displays *
      Set-SequencePanelVariable "  $VariableName`: $VariableQuery" $VarP
    }
    else {
      Set-SequencePanelVariable "  $VariableName`: ********" $VarP
    }
    $VarP++
    $SeqVariableHash[$TypePos].Add($VariableName, $VariableQuery)  # Add the pair Variable Name/Variable Value in the corresponding Variable Type hash
    if ($TypePos -eq 7) {
      # 7=Scheduler
      $Script:UseAScheduler = $True
      $Script:SchedulerVar = $VariableQuery
    }
  }

  return $SeqVariableHash  # Return the whole hash created

}


function Set-DefaultSettings {

  # Create all variables needed and set their default values

  New-Variable -Name CSVSeparator -Value ";" -Scope Script -Force
  New-Variable -Name CSVTempPath -Value "C:\Temp\HydraExport.csv" -Scope Script -Force
  New-Variable -Name XLSXTempPath -Value "C:\Temp\HydraExport.xlsx" -Scope Script -Force
  New-Variable -Name HTMLTempPath -Value "C:\Temp\HydraExport.html" -Scope Script -Force
  New-Variable -Name LogFilePath -Value "C:\Temp\Hydra.log" -Scope Script -Force
  New-Variable -Name CentralLogPath -Value "C:\Temp\" -Scope Script -Force
  New-Variable -Name LogFileEnabled -Value "True" -Scope Script -Force
  New-Variable -Name Colors -Value @{"OK" = "#FF90EE90" ; "BREAK" = "#FFADD8E6" ; "STOP" = "#FFF08080" ; "CANCELLED" = "#FFC0C0C0" } -Scope Script -Force
  New-Variable -Name DefaultThreads -Value "10" -Scope Script -Force
  New-Variable -Name DisplayWarning -Value $False -Scope Script -Force
  New-Variable -Name SCCM_ConfigMgrSiteServer -Value "" -Scope Script -Force
  New-Variable -Name SCCM_SiteCode -Value "" -Scope Script -Force
  New-Variable -Name NoSplashScreen -Value "False" -Scope Script -Force
  New-Variable -Name WelcomeScreen -Value "True" -Scope Script -Force
  New-Variable -Name DebugMode -Value 0 -Scope Script -Force
  New-Variable -Name ColorBackground -Value "#FF61B598" -Scope Script -Force
  New-Variable -Name ColorSequences -Value "#FFFFFFFF" -Scope Script -Force
  New-Variable -Name ColorSequencesRunning -Value "#FFFFFFE6" -Scope Script -Force
  New-Variable -Name LastDirObjects -Value "" -Scope Script -Force
  New-Variable -Name LastDirSequences -Value "" -Scope Script -Force
  New-Variable -Name LastDirExportGroup -Value "" -Scope Script -Force
  New-Variable -Name LastDirImportGroup -Value "" -Scope Script -Force
  New-Variable -Name ShowSearchBox -Value "True" -Scope Script -Force
  New-Variable -Name FileLoaded -Value "False" -Scope Script -Force
  New-Variable -Name CountriesList -Value "$HydraSettingsPath\Hydra_Countries.sccm" -Scope Script -Force
  New-Variable -Name ADQueriesList -Value "$HydraSettingsPath\Hydra_ADQueries.qry" -Scope Script -Force
  New-Variable -Name SequencesListPath -Value "$HydraSettingsPath\Hydra_Sequences.lst" -Scope Script -Force
  New-Variable -Name nbCheckedBoxes -Value 0 -Scope Script -Force
  New-Variable -Name LoadedFiles -Value "" -Scope Script -Force
  New-Variable -Name SequenceListExpanded -Value "True" -Scope Script -Force
  New-Variable -Name RowHeaderVisible -Value "False" -Scope Script -Force
  New-Variable -Name CheckBoxesKeepState -Value "False" -Scope Script -Force
  New-Variable -Name EMailSMTPServer -Value "" -Scope Script -Force
  New-Variable -Name EMailSendFrom -Value "" -Scope Script -Force
  New-Variable -Name EMailSendTo -Value "" -Scope Script -Force
  New-Variable -Name EMailReplyTo -Value "" -Scope Script -Force
  New-Variable -Name GrpCheckedOnWarning -Value "False" -Scope Script -Force
  New-Variable -Name GrpShowThreads -Value "True" -Scope Script -Force

  New-Variable -Name PosFormX -Value 100 -Scope Script -Force 
  New-Variable -Name PosFormY -Value 10 -Scope Script -Force
  New-Variable -Name PosFormW -Value 1150 -Scope Script -Force 
  New-Variable -Name PosFormH -Value 750 -Scope Script -Force 
  New-Variable -Name PosSplit1W -Value 1084 -Scope Script -Force 
  New-Variable -Name PosSplit1H -Value 688 -Scope Script -Force 
  New-Variable -Name PosSplit1D -Value 230 -Scope Script -Force
  New-Variable -Name PosSplit2W -Value 215 -Scope Script -Force 
  New-Variable -Name PosSplit2H -Value 688 -Scope Script -Force 
  New-Variable -Name PosSplit2D -Value 350 -Scope Script -Force
  New-Variable -Name SeqTreeTop -Value 35 -Scope Script -Force
  New-Variable -Name SeqTreeLeft -Value 10 -Scope Script -Force
  New-Variable -Name SeqTreeWidth -Value 215 -Scope Script -Force
  New-Variable -Name SeqTreeHeight -Value 310 -Scope Script -Force
  New-Variable -Name SeqPanelTop -Value 25 -Scope Script -Force
  New-Variable -Name SeqPanelLeft -Value 10 -Scope Script -Force
  New-Variable -Name SeqPanelWidth -Value 215 -Scope Script -Force
  New-Variable -Name SeqPanelHeight -Value 290 -Scope Script -Force
  New-Variable -Name PanelBottomTop -Value 600 -Scope Script -Force
  New-Variable -Name PanelBottomWidth -Value 840 -Scope Script -Force
  New-Variable -Name DataGridTabControlWidth -Value 825 -Scope Script -Force
  New-Variable -Name DataGridTabControlHeight -Value 580 -Scope Script -Force
  New-Variable -Name TabLook -Value "1" -Scope Script -Force

  New-Variable -Name TimerIntervalDefault -Value 1000 -Scope Script -Force
  New-Variable -Name TimerInterval -Value $TimerIntervalDefault -Scope Script -Force
  New-Variable -Name SendMail -Value $False -Scope Script -Force
  New-Variable -Name TimerSet -Value $False -Scope Script -Force
  New-Variable -Name NewGroupName -Value 0 -Scope Script -Force
  New-Variable -Name RemovedFromSeqID -Value 0 -Scope Script -Force
  New-Variable -Name SelectionChanged -Value $False -Scope Script -Force
  New-Variable -Name VariablesQuery -Value $Null -Scope Script -Force

  $Script:TabColorPalette = [ordered]@{"DodgerBlue" = "LightBlue"; "Crimson" = "LightCoral"; "SpringGreen" = "LightGreen"; "Yellow" = "Khaki"; "DarkGray" = "LightGray"; "DarkOrchid" = "Orchid"; "Sienna" = "Chocolate" }
  
  $Script:Sequences = , @()
  
  $Script:RunspaceCollection = , @() 
  $Script:RunspacePool = , @()
  $Script:SequencesToParse = New-Object System.Collections.ArrayList
  $Script:OutputDataGridSequence = , @()
  $Script:JobNb = , @()
  $Script:ConcurrentJobs = , @()
  $Script:SequenceAssigned = $False
  $Script:SequenceTabIndex = , @()
  $Script:GroupsUsed = @()
  $Script:GroupsRunning = @()

}


function Set-GroupMaxThreads {

  # Modify the number of threads of a group

  $NewMaxThread = (Read-InputBoxDialog "Max. Threads" "Enter the new maximum threads for the Group" 1) -as [int]

  if (($NewMaxThread -lt 1) -or ($NewMaxThread -gt 1000)) {
    return
  }
  $SelectedRows = @($OutputDataGrid.Rows[0..$($OutputDataGrid.Rows.Count - 2)] | where { $_.Cells[0].Tag.GroupID -eq $GroupsFoundForMaxThreads })  # Select all rows belonging to the group to modify
  foreach ($row in $SelectedRows) {  
    $Script:Sequences[$row.Cells[8].Value].MaxThreads = $NewMaxThread  # Modify the MaxThreads value for the Sequence
    $row.Cells[9].Value = "$GroupsFoundForMaxThreads"  # Set the Text to display in the column Group as the name of the Group
    if ($GrpShowThreads -eq "True") { $row.Cells[9].Value += " ($NewMaxThread)" }  # Add the value of the Max Threads in the column Group if the user's settings is set
  }

}


function Set-ObjectsState {
  
  # Print the number of objects in the current grid as well as the number of objects selected

  $ObjectsLabel.Text = "Total Objects: $($OutputDataGrid.RowCount-1) ,   Selected: $nbCheckedBoxes"

}


function Set-ObjectsToGroup($GroupToAdd) {

  # Add objects to an existing group

  $RowRef = $Null

  for ($i = 0; $i -le $OutputDataGrid.RowCount - 1; $i++) {
    # Parse the grid and search for the 1st object in the group to add
    if ($OutputDataGrid.Rows[$i].Cells[0].Tag.GroupID -eq $GroupToAdd) {
      $RowRef = $OutputDataGrid.Rows[$i]  # Set $RowRef as row reference of the group
      break
    }
  }

  if ($RowRef -eq $Null) { return } 

  $SeqValue = $RowRef.Cells[8].Value  # Determine the Sequence ID, the number of threads and the name of the sequence
  $MaxThreadValue = $Script:Sequences[$RowRef.Cells[8].Value].MaxThreads
  $SeqName = $Sequences[$SeqValue].SequenceLabel

  $DataGridViewCellStyleBold = New-Object System.Windows.Forms.DataGridViewCellStyle
  $DataGridViewCellStyleBold.Alignment = 16
  $DataGridViewCellStyleBold.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [Drawing.FontStyle]::"Bold", 3, 0)
  $DataGridViewCellStyleBold.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
  $DataGridViewCellStyleBold.SelectionBackColor = [System.Drawing.Color]::FromArgb(255, 51, 153, 255)
  $DataGridViewCellStyleBold.SelectionForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 255)

  $DataGridViewCellStyleRegular = New-Object System.Windows.Forms.DataGridViewCellStyle
  $DataGridViewCellStyleRegular.Alignment = 16
  $DataGridViewCellStyleRegular.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [Drawing.FontStyle]::"Regular", 3, 0)
  $DataGridViewCellStyleRegular.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
  $DataGridViewCellStyleRegular.SelectionBackColor = [System.Drawing.Color]::FromArgb(255, 51, 153, 255)
  $DataGridViewCellStyleRegular.SelectionForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 255)

  foreach ($item in $OutputDataGrid.SelectedCells) {
    # Set the Group value to all objects selected
    $RowIndex = $item.RowIndex
    Set-CellValue $GridIndex $RowIndex 0 "Sequence assigned: $SeqName" "Pending" 0 "#" "#" $SeqValue  # Set the object as pending
    $OutputDataGrid.Rows[$RowIndex].Cells[0].ReadOnly = $True
    $OutputDataGrid.Rows[$RowIndex].Cells[0].Tag.GroupID = $GroupToAdd  # Tag the object with the Group name
    $OutputDataGrid.Rows[$RowIndex].DefaultCellStyle.BackColor = "White"
    $OutputDataGrid.Rows[$RowIndex].Cells[0].Style = $DataGridViewCellStyleBold
    $OutputDataGrid.Rows[$RowIndex].Cells[2].Style = $DataGridViewCellStyleRegular
    $OutputDataGrid.Rows[$RowIndex].Cells[9].Value = "$GroupToAdd"  # Set the text to print in the column Group
    if ($GrpShowThreads -eq "True") { $OutputDataGrid.Rows[$RowIndex].Cells[9].Value += " ($MaxThreadValue)" }
    $Script:OutputDataGridSequence[$SeqValue] += $OutputDataGrid.Rows[$RowIndex]  # Add the current row to the OutputDataGridSequence relative of the Sequence ID
  }
  
}


function Set-RecreateGroups {

  # Parse the grid and recreate the groups after deleting operations have been performed

  # Search for all Sequence ID's assigned to groups in the grid
  $SeqFound = for ($i = 0; $i -lt $DataGridTabControl.TabCount; $i++) { 
    $TabGridID = $DataGridTabControl.TabPages[$i].Tag.TabPageIndex  # Get the Grid ID
    $OutputDataGridTab[$TabGridID].Rows | where { ($_.Cells[0] | select -ExpandProperty Tag) | where { $_.GroupID -ne "0" } } | foreach { $_.Cells[8].Value } | select -Unique | Sort-Object
  }

  if (@($SeqFound).Count -eq 0) {
    # No Sequence ID relative to any group found: no group defined (anymore)
    $Script:GroupsUsed = @()
    return 
  }

  foreach ($Seq in $SeqFound) { $Script:OutputDataGridSequence[$Seq] = @() }  # Reset all OutputDataGridSequence relative to the Sequence ID's found

  for ($i = 0; $i -lt $OutputDataGrid.RowCount - 1; $i++) {
    # Parse the grid and refill the OutputDataGridSequence arrays
    if ($OutputDataGrid.Rows[$i].Cells[0].Tag.GroupID -ne "0") {
      $SeqId = $OutputDataGrid.Rows[$i].Cells[8].Value
      $Script:OutputDataGridSequence[$SeqId] += $OutputDataGrid.Rows[$i]
    }
  }

  $Script:GroupsUsed = @()
  # Parse all tabs, get the name of all the Groups and store them in $GroupsUsed
  $Script:GroupsUsed = for ($i = 0; $i -lt $DataGridTabControl.TabCount; $i++) { 
    $TabGridID = $DataGridTabControl.TabPages[$i].Tag.TabPageIndex  # Get the Grid ID
    ($OutputDataGridTab[$TabGridID].Rows.Cells | where { ($_.ColumnIndex -eq 0) } | select -ExpandProperty Tag) | where { $_.GroupID -ne 0 } | select -ExpandProperty GroupID -Unique | Sort-Object 
  }

}


function Set-ReloadSequenceList {

  # Reload the Sequence List for potential changes or correction

  $SequencesTreeView.Nodes.Clear()
  SequencesTreeView_GetSequenceList $True
  if ($FormSettingsSequenceExpandedRadioButton.Checked) {
    # Collapse or expand depending on the user's settings
    $SequencesTreeView.ExpandAll()
  }
  else { 
    $SequencesTreeView.CollapseAll()
  }

}


function Set-RightClick_CheckObject($Check) {

  # Check or uncheck the objects selected

  foreach ($RowIndex in $OutputDataGrid.SelectedCells.RowIndex) { $OutputDataGrid.Rows[$RowIndex].Cells[7].Value = $Check }

  Get-CountCheckboxes

}


function Set-RightClick_RemoveSelectionFromFiles ($FileList, $DeleteRows) {

  # Remove the selected objects from the files they were loaded from

  $IsRunningSeqSelected = @($OutputDataGrid.SelectedRows | where { $_.Cells[4].Value -gt 0 }).Count
  if ($IsRunningSeqSelected -gt 0) { return }  # If some objects are running, exits

  foreach ($File in $FileList) {
    # Generate the objects to remove for each file 
    $ObjectsRemaining = @()
    foreach ($SelectedItem in $OutputDataGrid.SelectedCells) {
      if ($OutputDataGrid.Rows[$SelectedItem.RowIndex].Cells[6].Value -eq $File) {
        # The file name matches: adds the object in the array
        $ObjectsRemaining += $SelectedItem.Value
      }
    }
    $NewText = Select-String -Path $File -Pattern $ObjectsRemaining -NotMatch | Select-Object -ExpandProperty 'Line'  # Recreate the list of objects removing the ones stored in the array
    $NewText | Set-Content -Path $File  # Recreate the objects file with the new content
  }

  if ($DeleteRows) {
    # Delete the objects from the grid too if needed
    Set-RightClick_SetNewSelectionFromGrid $False
  }

}


function Set-RightClick_SetNewSelectionFromGrid ($Action) {

  # Recreate the grid with or without the selected objects

  $IsRunningSeqSelected = @($OutputDataGrid.SelectedRows | where { $_.Cells[4].Value -gt 0 }).Count
  if ($IsRunningSeqSelected -gt 0) { return }  # If some objects are running, exits

  $NewSelection = @()
  $NewSelectionCellStyle = @()
  for ($i = 0; $i -lt $OutputDataGrid.RowCount - 1; $i++) {
    # Create a new selection based on $Action
    if ($OutputDataGrid.Rows[$i].Cells[0].Selected -eq $Action) {
      # If $Action=True, add the selected objects, if $Action=False, add the non-selected objects 
      $NewSelection += , @($OutputDataGrid.Rows[$i])
      $NewSelectionCellStyle += , @($OutputDataGrid.Rows[$i].Cells[2].Style)
    }
  }

  $SeqFound = $NewSelection | foreach { $_.Cells[8].Value } | select -Unique | Sort-Object

  foreach ($Seq in $SeqFound) { $Script:OutputDataGridSequence[$Seq] = @() }  # Reset the OutputDataGridSequence for all sequences found in the new selection

  $OutputDataGrid.Rows.Clear()

  $DataGridViewCellStyleBold = New-Object System.Windows.Forms.DataGridViewCellStyle
  $DataGridViewCellStyleBold.Alignment = 16
  $DataGridViewCellStyleBold.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [Drawing.FontStyle]::"Bold", 3, 0)
  $DataGridViewCellStyleBold.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
  $DataGridViewCellStyleBold.SelectionBackColor = [System.Drawing.Color]::FromArgb(255, 51, 153, 255)
  $DataGridViewCellStyleBold.SelectionForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 255)

  for ($i = 0; $i -lt $NewSelection.Count; $i++) {
    # Recreate the new grid based on the newly created selection
    $OutputDataGrid.Rows.Add($NewSelection[$i].Cells.Value)  # Clone the row and add the Tags
    $OutputDataGrid.Rows[$i].Cells[0].Tag = $NewSelection[$i].Cells[0].Tag
    if ($OutputDataGrid.Rows[$i].Cells[0].Tag.GroupID -ne "0") {
      # Group found
      $OutputDataGrid.Rows[$i].Cells[0].Style = $DataGridViewCellStyleBold
      $OutputDataGrid.Rows[$i].Cells[0].ReadOnly = $True
    }
    $SeqId = $OutputDataGrid.Rows[$i].Cells[8].Value
    $Script:OutputDataGridSequence[$SeqId] += $OutputDataGrid.Rows[$i]  # Add the row to the OutputDataGridSequence of the sequence ID
    if ($OutputDataGrid.Rows[$i].Cells[4].Value -lt 0) {
      # A sequence ran already: don't reset the color
      $OutputDataGrid.Rows[$i].Cells[2].Style = $($NewSelectionCellStyle[$i])
      $OutputDataGrid.Rows[$i].DefaultCellStyle.BackColor = $OutputDataGrid.Rows[$i].Cells[5].Value
    }
  }

  Set-RecreateGroups

  Get-CountCheckboxes

}


function Set-RightClick_SetNewSelectionFromState($State, $DeleteRows) {

  # Select or remove objects based on a state

  $IsRunningSeqSelected = @($OutputDataGrid.SelectedRows | where { $_.Cells[4].Value -gt 0 }).Count
  if ($IsRunningSeqSelected -gt 0) { return }  # If some objects are running, exits

  $OutputDataGrid.ClearSelection()
  for ($i = 0; $i -lt $OutputDataGrid.RowCount - 1; $i++) {
    # Select the objects matching the state
    if ($OutputDataGrid.Rows[$i].Cells[3].Value -eq $State) {  
      $OutputDataGrid.Rows[$i].Cells[0].Selected = $True
    }
  }

  if ($DeleteRows) {
    # Delete the objects selected if necessary
    Set-RightClick_SetNewSelectionFromGrid $False
  }
  else {
    Set-RecreateGroups
    Get-CountCheckboxes
  }

}


function Set-RightClick_ShowProtocol {
  
  # Display the steps protocol of the selected objects
   
  $NewText = @()  # Protocol to display, first set to nothing
  $ProtocolFound = $False
  $RowsToKeep = $OutputDataGrid.Rows[ ($OutputDataGrid.SelectedCells | select -ExpandProperty RowIndex) ]  # Get the row indexes to treat
  foreach ($Row in $RowsToKeep) { 
    $NewText += "$(($Row.Cells[0].Tag.StepProtocol) -join "`r`n")`r`n"  # Add the protocol of the object (Cells[0].Tag.StepProtocol) to the protocol to display
    if ($Row.Cells[0].Tag.StepProtocol -ne $Null) { $ProtocolFound = $True }
  }
  
  if ($ProtocolFound -eq $False) {
    $NewText = "No protocol found"  
  }

  $OutputDataGridContextMenuObjectProtocol.DropDownItems[0].Text = $NewText  # Add the text to the right click menu

}


function Set-RunspaceToObjects {

  # Allocate a new Runspace Pool to any sequence to start

  # Search all sequences existing in the grid, for checked objects 
  $RowChecked = $OutputDataGrid.Rows.Cells | where { $_.ColumnIndex -eq 7 } | where { $_.Value -eq $True } | select -ExpandProperty RowIndex -Unique
  $RowsSeqIDs = foreach ($row in $RowChecked) { $OutputDataGrid.Rows[$row].Cells[8] }
  $SequenceIndex = $RowsSeqIDs | where { $_.Value -ne 0 } | select -ExpandProperty Value -Unique

  foreach ($Index in $SequenceIndex) {
    # Prompt for security for the groups
    if ($Sequences[$Index].BelongsToGroup) {
      if ($Sequences[$Index].SecurityCode -ne "") {
        $CodePrompt = Read-SecurityCode $($Sequences[$Index].SecurityCode) $($Sequences[$Index].SequenceLabel)
        if ($CodePrompt -ne "OK") {
          $SequenceIndex = $SequenceIndex | Where-Object { $_ –ne $Index }  # Sequence cancelled for the objects of this group, remove the index
          continue
        }
      }
      elseif ($Sequences[$Index].DisplayWarning -ne $False) {
        # The sequence has the Warning option
        $ReallyDeploy = Read-StartSequence $($Sequences[$Index].SequenceLabel)
        if ($ReallyDeploy -ne "OK") {
          $SequenceIndex = $SequenceIndex | Where-Object { $_ –ne $Index }  # Sequence cancelled for the objects of this group, remove the index
          continue
        }
      }
    }
  }

  if ($SequenceIndex -eq $null) { return }

  # Search for the sequences not currently running 
  $MissingIndex = $SequenceIndex | where { -not ($SequencesToParse -contains $_) }
  $CentralLogFilePath = $FormSettingsPathsValue[4] + "\" + [Environment]::UserName + "\"

  foreach ($Index in $MissingIndex) {
    # Create the Runspace Pool for each sequence 

    if ($Sequences[$Index].SequenceSchedulerExpired) {
      $OutputDataGridSequence[$Index].Cells | where { $_.ColumnIndex -eq 7 } | foreach { $_.Value = $False }  # Uncheck all the objects of the sequence to skip it
      $OutputDataGridSequence[$Index].Cells | where { $_.ColumnIndex -eq 3 } | foreach { $_.Value = "Timer Expired" }  # Uncheck all the objects of the sequence to skip it
      $OutputDataGrid.RefreshEdit()
      continue
    }
    
    if ($Sequences[$Index].SequenceScheduler -ne 0) {
      # A timer is defined
      $TimeDiff = New-TimeSpan $(Get-Date) $Sequences[$Index].SequenceScheduler
      $TimeDiffFormated = '{0:00}:{1:00}:{2:00}' -f $TimeDiff.Hours, $TimeDiff.Minutes, $TimeDiff.Seconds
      if ($TimeDiff.TotalSeconds -le 1) {
        # The timer has expired
        $Script:Sequences[$Index].SequenceScheduler = 0 
        $Script:Sequences[$Index].SequenceSchedulerExpired = $True
        $OutputDataGridSequence[$Index].Cells | where { $_.ColumnIndex -eq 7 } | foreach { $_.Value = $False }  # Uncheck all the objects of the sequence to skip it
        $OutputDataGridSequence[$Index].Cells | where { $_.ColumnIndex -eq 3 } | foreach { $_.Value = "Timer Expired" }  # Uncheck all the objects of the sequence to skip it
        $OutputDataGrid.RefreshEdit()
        continue
      }
    }

    $Script:RunspaceCollection[$Index] = , @()
    $SessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()  # Create a InitialSessionState and add default variables
    $SessionState.Variables.Add( (New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry("MyScriptInvocation", $MyScriptInvocation, $null)) )
    $SessionState.Variables.Add( (New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry("SequencePath", $SequencePath, $null)) )
    $SessionState.Variables.Add( (New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry("SequenceFullPath", $SequenceAbsolutePath, $null)) )
    $SessionState.Variables.Add( (New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry("CentralLogPath", $CentralLogFilePath, $null)) )
    
    foreach ($Module in $Sequences[$Index].ScriptBlockModule) {
      # Add modules to the InitialSessionState if some have been declared in the sequence.xml file
      if ($Module.Type -eq "ImportPSSnapIn") { [void]$SessionState.ImportPSSnapIn($($Module.Name), [ref]$null) }
      if ($Module.Type -eq "ImportPSModulesFromPath") { [void]$SessionState.ImportPSModulesFromPath($($Module.Name)) }
      if ($Module.Type -eq "ImportPSModule") { [void]$SessionState.ImportPSModule($($Module.Name)) }
    }

    for ($i = 1; $i -le $nbVariableTypes; $i++) {
      # Add variables names and their values to the InitialSessionState if some have been declared in the sequence.xml file
      $Sequences[$Index].ScriptBlockVariable[$i].keys | foreach { 
        $SessionState.Variables.Add( (New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry($_, $Sequences[$Index].ScriptBlockVariable[$i].Item($_), $null)) )
      }
    }
    $Script:RunspacePool[$Index] = [RunspaceFactory]::CreateRunspacePool(1, $Sequences[$Index].MaxThreads, $SessionState, $Host)  # Create and open the RunspacePool with the right parameters
    $Script:RunspacePool[$Index].Open()
    if ($SequencesToParse -notcontains $Index) { $Script:SequencesToParse.Add($Index) }  # Add the sequence to the list of sequence to parse at each timer interval
    $Script:JobNb[$Index] = 0
    $Script:ConcurrentJobs[$Index] = 1
  } 

}


function Set-SearchBox {

  # Show or Hide the Sequence Search Box

  if ($ShowSearchBox -eq "True") { $SequencesTreeViewTopPosition = 20 } else { $SequencesTreeViewTopPosition = 0 }
  $SequencesTreeView.Top = [int]$SeqTreeTop + $SequencesTreeViewTopPosition 
  $SequencesTreeView.Height = $SeqTreeHeight - $SequencesTreeViewTopPosition
  $SearchTreeTextBox.Visible = ($ShowSearchBox -eq "True")

}


function Set-SelectGroup($Group, $Select) {

  # Select or deselect objects of groups

  $OutputDataGrid.ClearSelection()
  if ($Group -eq "All Groups") {
    $GroupsToSelect = @(($OutputDataGrid.Rows.Cells | where { ($_.ColumnIndex -eq 0) } | select -ExpandProperty Tag) | where { $_.GroupID -ne 0 } | select -ExpandProperty GroupID -Unique | Sort-Object)  # Enumerate the groups in the grid
  }
  else {
    $GroupsToSelect = @($Group)
  }
  
  foreach ($row in $OutputDataGrid.Rows) {
    # Parse the grid and select the matching objects
    if (($row.Cells[0].Tag.GroupID) -in $GroupsToSelect) { $row.Cells[7].Value = $Select }
  }
  
  $OutputDataGrid.EndEdit()

  Get-CountCheckboxes

}


function Set-SequenceFinished {

  # Reset the components when all Sequences have finished

  $Timer.Enabled = $False  # Stop the Timer
  $Script:SequenceRunning = $False
  $ObjectsLabel.Text = ""
  Send-MailLog  # Send a Mail with the results if it was defined in some Sequences
  Write-Log
  foreach ($Item in $SequencesToParse) {
    # Re-enable the Checkboxes
    foreach ($Row in $OutputDataGridSequence[$Item]) { $Row.Cells[7].ReadOnly = $False }
  }
  Reset-Runspaces
  Reset-SequenceArrays
  Get-CountCheckboxes

  # Disable some menu entries
  ($MenuMain.Items | where { $_.Text -eq "Cancel" }).Enabled = $False
  ($MenuToolStrip.Items | where { $_.ToolTipText -eq "Cancel All" }).Enabled = $False
  ($MenuMain.Items.DropDown.Items | where { $_.Text -eq "Reset All Objects" }).Enabled = $True
  ($MenuMain.Items.DropDown.Items | where { $_.Text -eq "Clear Grid" }).Enabled = $True

}


function Set-SequencePanelCheckbox($Text, $SeqPos) {

  # Display a Checkbox for the Sequence Steps, with the Text and Position given as parameters

  $SequencePanelTempCheckbox = New-Object System.Windows.Forms.Checkbox
  $Pos_Y = $(43 * $SeqPos + 45)  # Calculate the vertical position based on the Position parameter
  $SequencePanelTempCheckbox.Location = New-Object System.Drawing.Size(15, $Pos_Y)
  $SequencePanelTempCheckbox.Checked = $True
  $SequencePanelTempCheckbox.ForeColor = 'Black'
  $SequencePanelTempCheckbox.AutoSize = $True
  $SequencePanelTempCheckbox.MaximumSize = New-Object System.Drawing.Size(500, 15)
  $SequencePanelTempCheckbox.Font = New-Object Drawing.Font("Microsoft Sans Serif", 8, [Drawing.FontStyle]::"Regular")
  $SequencePanelTempCheckbox.ForeColor = [Drawing.Color]::"Black"
  $SequencePanelTempCheckbox.Text = $Text
  $Script:SequencePanelCheckbox += $SequencePanelTempCheckbox
  $SequenceTasksPanel.Controls.Add($SequencePanelCheckbox[$SeqPos])
  
}


function Set-SequencePanelLabel($Text, $Style, $Color, $SeqPos) {

  # Display a Label for the Sequence Steps, with the Text, Style, Color and Position given as parameters

  $SequencePanelTempLabel = New-Object System.Windows.Forms.Label
  $SequencePanelTempLabel.Text = $Text
  $Pos_Y = $(43 * $SeqPos + 62)  # Calculate the vertical position based on the Position parameter
  $SequencePanelTempLabel.Location = New-Object System.Drawing.Size(15, $Pos_Y)
  $SequencePanelTempLabel.AutoSize = $True
  $SequencePanelTempLabel.MaximumSize = New-Object System.Drawing.Size(500, 15)
  $SequencePanelTempLabel.Font = New-Object Drawing.Font("Microsoft Sans Serif", 8, [Drawing.FontStyle]::$Style)
  $SequencePanelTempLabel.ForeColor = [Drawing.Color]::$Color
  $Script:SequencePanelLabel += $SequencePanelTempLabel
  $SequenceTasksPanel.Controls.Add($SequencePanelLabel[$SeqPos])
  
}


function Set-SequencePanelTitle($Text, $Color) {

  # Display the Title of the Sequence, with the Text and Color given as parameters

  $SequencePanelTitleLabel = New-Object System.Windows.Forms.Label
  $SequencePanelTitleLabel.Text = $Text
  $SequencePanelTitleLabel.Location = New-Object System.Drawing.Size(10, 10)
  $SequencePanelTitleLabel.AutoSize = $True
  $SequencePanelTitleLabel.MaximumSize = New-Object System.Drawing.Size(500, 15)
  $SequencePanelTitleLabel.Font = New-Object Drawing.Font("Tahoma", 9, [Drawing.FontStyle]::"Underline")
  $SequencePanelTitleLabel.ForeColor = [Drawing.Color]::$Color
  $SequenceTasksPanel.Controls.Add($SequencePanelTitleLabel)

}


function Set-SequencePanelVariable($Text, $VarPos) {

  # Display the value of a Variable, with the Text and Position given as parameters

  $SequencePanelTempVariable = New-Object System.Windows.Forms.Label
  $SequencePanelTempVariable.Text = $Text -replace "`n", "|"  # Display every objects on one line only
  $Pos_Y = 44 * $MaxSteps + 16 * $VarPos + 48  # Calculate the vertical position based on the Position parameter
  $SequencePanelTempVariable.Location = New-Object System.Drawing.Size(15, $Pos_Y)
  $SequencePanelTempVariable.AutoSize = $True
  $SequencePanelTempVariable.MaximumSize = New-Object System.Drawing.Size(500, 15)
  $SequencePanelTempVariable.Font = New-Object Drawing.Font("Microsoft Sans Serif", 8, [Drawing.FontStyle]::"Regular")
  $SequencePanelTempVariable.ForeColor = [Drawing.Color]::"Blue"
  $Script:SequencePanelVariable += $SequencePanelTempVariable
  $SequenceTasksPanel.Controls.Add($SequencePanelVariable[$VarPos])
  
}


function Set-SettingsSubMenu($SubMenu) {

  # Sub-function to display the correct groupboxes and elements in the Settings menu

  $FormSettingsPathsGroupBox.Visible = $False
  $FormSettingsColorsGroupBox.Visible = $False
  $FormSettingsColorsGUIGroupBox.Visible = $False
  $FormSettingsSequenceSearchGroupBox.Visible = $False
  $FormSettingsMiscGroupBox.Visible = $False
  $FormSettingsSequenceGroupBox.Visible = $False
  $FormSettingsRowHeaderGroupBox.Visible = $False
  $FormSettingsCheckBoxesGroupBox.Visible = $False
  $FormSettingsMailGroupBox.Visible = $False
  $FormSettingsGroupsWarningGroupBox.Visible = $False
  $FormSettingsGroupsThreadsGroupBox.Visible = $False
  $FormSettingsTabsLookGroupBox.Visible = $False

  switch ($SubMenu) {
    "Paths" {
      $FormSettingsPathsGroupBox.Visible = $True
    }
    "Colors" {
      $FormSettingsColorsGroupBox.Visible = $True
      $FormSettingsColorsGUIGroupBox.Visible = $True
    }
    "Misc" {
      $FormSettingsSequenceSearchGroupBox.Visible = $True
      $FormSettingsMiscGroupBox.Visible = $True
      $FormSettingsSequenceGroupBox.Visible = $True
      $FormSettingsRowHeaderGroupBox.Visible = $True
      $FormSettingsCheckBoxesGroupBox.Visible = $True
    }
    "Mail" {
      $FormSettingsMailGroupBox.Visible = $True
    }
    "Groups" {
      $FormSettingsGroupsWarningGroupBox.Visible = $True
      $FormSettingsGroupsThreadsGroupBox.Visible = $True
    }
    "Tabs" {
      $FormSettingsTabsLookGroupBox.Visible = $True
    }
  }

}


function Set-StartupSettings($Invocation) {

  # Startup checks and Settings

  if ($Host.Version.Major -lt 4) {
    # The Powershell version is lower than 4: exit
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void][System.Windows.Forms.MessageBox]::Show("You need Powershell 4 or higher to run this version of Hydra.`n`r`n`rPlease install a newer version`r`n" , "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
  }

  if (!(Test-Path 'HKCU:\Software\Hydra')) {
    # Create the Hydra registry structure if it's missing
    New-Item -Path 'HKCU:\Software' -Name Hydra | Out-Null
  }
  else {
    if (!(Test-Path 'HKCU:\Software\Hydra\3')) {
      # Clean-up the Hydra3 registry
      Copy-Item -Path 'HKCU:\Software\Hydra' -Destination 'HKCU:\Software\Hydra3' | Out-Null
      Get-Item 'HKCU:\Software\Hydra' | Remove-Item -Recurse -Force
      New-Item -Path 'HKCU:\Software' -Name Hydra | Out-Null
      Copy-Item -Path 'HKCU:\Software\Hydra3' -Destination 'HKCU:\Software\Hydra\3' -Force | Out-Null
      Get-Item 'HKCU:\Software\Hydra3' | Remove-Item -Recurse -Force
    }
  }

  if (!(Test-Path 'HKCU:\Software\Hydra\5')) {
    # Copy the settings of Hydra4 to Hydra5
    if (Test-Path 'HKCU:\Software\Hydra4') {
      Copy-Item -Path 'HKCU:\Software\Hydra4' -Destination 'HKCU:\Software\Hydra\5' -Force | Out-Null
      $ToDelete = "DataGridHeight", "DataGridWidth", "PanelBottomTop", "PanelBottomWidth", "PosFormH", "PosFormW", "PosFormX", "PosFormY", "PosSplit1D", "PosSplit1H", "PosSplit1W", "PosSplit2D", "PosSplit2H", "PosSplit2W", "SeqPanelHeight", "SeqPanelLeft", "SeqPanelTop", "SeqPanelWidth", "SeqTreeHeight", "SeqTreeLeft", "SeqTreeTop", "SeqTreeWidth", "WelcomeScreen"
      foreach ($item in $ToDelete) {
        # Remove obsolete Hydra4 variables
        Remove-ItemProperty HKCU:\Software\Hydra\5 -Name $item -ErrorAction SilentlyContinue
      }
    }
    else {
      New-Item -Path 'HKCU:\Software\Hydra' -Name 5 | Out-Null
    }
  }

  #Define global varibales

  New-Variable -Name PSScriptName -Value $($Invocation.MyCommand.Name) -Scope Script -Force
  New-Variable -Name HydraBinPath -Value $(Split-Path $Invocation.InvocationName) -Scope Script -Force
  New-Variable -Name MyScriptInvocation -Value $Invocation -Scope Script -Force
  New-Variable -Name HydraSettingsPath -Value "$PSScriptRoot\Settings" -Scope Script -Force
  New-Variable -Name HydraGUIPath -Value "$PSScriptRoot\GUI" -Scope Script -Force
  New-Variable -Name HydraDocsPath -Value "$PSScriptRoot\Docs" -Scope Script -Force
  New-Variable -Name SequenceName -Value "" -Scope Script -Force
  New-Variable -Name ResetSettings -Value $False -Scope Script -Force
  New-Variable -Name SequenceRunning -Value $False -Scope Script -Force

  Set-DefaultSettings  # Set the default variables settings
  Get-RegistrySettings  # Get registry values and replace the default ones defined the step before

  if ($SequencesListParam -ne $Null) {
    # Hydra has been started with a Sequences List as parameter
    $Script:SequencesListPath = $SequencesListParam  # Set the global variable SequencesListPath with this parameter
  }

  Set-Location $PSScriptRoot  # Set the current directory to the Hydra path

  if ( (Test-Path $SequencesListPath) -eq $False) {
    # No Sequences List found: exit
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void][System.Windows.Forms.MessageBox]::Show("Unable to find the Sequences List" , "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
  }

  # Load the Sequence Variables Types
  Set-StartupSettings_SeqVariables

  # Show or Hide the Powershell console
  Add-Type -Name Window -Namespace Console -MemberDefinition '
  [DllImport("Kernel32.dll")]
  public static extern IntPtr GetConsoleWindow(); 

  [DllImport("user32.dll")]
  public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
  $consolePtr = [Console.Window]::GetConsoleWindow()
  [void] [Console.Window]::ShowWindow($consolePtr, $DebugMode)  # 0 to make the Powershell console invisible, 5 to make the Powershell console visible 

}


function Set-StartupSettings_SeqVariables {

  # Define the variables types

  $Script:nbVariableTypes = 10
  $Script:SeqVariableHash = @(0) * ($nbVariableTypes + 1)
  $Script:VariableCommand = @(0) * ($nbVariableTypes + 1)

  # Define the names of the variables that can be used in a .sequence.xml
  $Script:VariableTypes = @("", "inputbox", "multilineinputbox", "selectfile", "selectfolder", "combobox", "multicheckbox", "scheduler", "credentials", "secretinputbox", "credentialbox")

  # Create a script block based on a command to any of the variable types. It will pass $SeqVariables defined by the user
  $Script:VariableCommand[1] = @'
    switch ($SeqVariables.Split(';').Count) {
      0 { Read-InputBoxDialog }
      1 { Read-InputBoxDialog $SeqVariables.Split(';')[0] }
      2 { Read-InputBoxDialog $SeqVariables.Split(';')[0] $SeqVariables.Split(';')[1] }
      {$_ -ge 3} { Read-InputBoxDialog $SeqVariables.Split(';')[0] $SeqVariables.Split(';')[1] $SeqVariables.Split(';')[2] }

    }
'@
  $Script:VariableCommand[2] = @'
    switch ($SeqVariables.Split(';').Count) {
      0 { Read-MultiLineInputBoxDialog }
      1 { Read-MultiLineInputBoxDialog $SeqVariables.Split(';')[0] }
      2 { Read-MultiLineInputBoxDialog $SeqVariables.Split(';')[0] $SeqVariables.Split(';')[1] }
      {$_ -ge 3} { Read-MultiLineInputBoxDialog $SeqVariables.Split(';')[0] $SeqVariables.Split(';')[1] $SeqVariables.Split(';')[2] }
    }
'@
  $Script:VariableCommand[3] = @'
    switch ($SeqVariables.Split(';').Count) {
      0 { Read-OpenFileDialog }
      1 { Read-OpenFileDialog $SeqVariables.Split(';')[0] }
      2 { Read-OpenFileDialog $SeqVariables.Split(';')[0] $SeqVariables.Split(';')[1] }
      {$_ -ge 3} { $LastParamPos=($SeqVariables.Split(';')[0]).Length+($SeqVariables.Split(';')[1]).Length+2
          Read-OpenFileDialog $SeqVariables.Split(';')[0] $SeqVariables.Split(';')[1] $SeqVariables.Substring($LastParamPos, $SeqVariables.Length-$LastParamPos)
        }
    }
'@
  $Script:VariableCommand[4] = @'
    fBrowse-Folder_Modern($SeqVariables)  
'@
  $Script:VariableCommand[5] = @'
    switch ($SeqVariables.Split(';').Count) {
      0 { Read-ComboBoxDialog }
      1 { Read-ComboBoxDialog $SeqVariables.Split(';')[0] }
      2 { Read-ComboBoxDialog $SeqVariables.Split(';')[0] $SeqVariables.Split(';')[1] }
      {$_ -ge 3} { Read-ComboBoxDialog $SeqVariables.Split(';')[0] $SeqVariables.Split(';')[1] $SeqVariables.Split(';')[2] }
    }
'@
  $Script:VariableCommand[6] = @'
     switch ($SeqVariables.Split(';').Count) {
       0 { Read-MultiCheckboxList }
       1 { Read-MultiCheckboxList $SeqVariables.Split(';')[0] }
       2 { Read-MultiCheckboxList $SeqVariables.Split(';')[0] $SeqVariables.Split(';')[1] }
       {$_ -ge 3} { Read-MultiCheckboxList $SeqVariables.Split(';')[0] $SeqVariables.Split(';')[1] $SeqVariables.Split(';')[2] }
     } 
'@
  $Script:VariableCommand[7] = @'
    Read-DateTimePicker $SeqVariables
'@
  $Script:VariableCommand[8] = @'
    Read-Credentials $SeqVariables
'@
  $Script:VariableCommand[9] = @'
    switch ($SeqVariables.Split(';').Count) {
      0 { Read-InputDialogBoxSecret }
      1 { Read-InputDialogBoxSecret $SeqVariables.Split(';')[0] }
      {$_ -ge 2} { Read-InputDialogBoxSecret $SeqVariables.Split(';')[0] $SeqVariables.Split(';')[1] }
    }
'@
  $Script:VariableCommand[10] = @'
    switch ($SeqVariables.Split(';').Count) {
      0 { Read-Credentialbox }
      1 { Read-Credentialbox $SeqVariables.Split(';')[0] }
      {$_ -ge 2} { Read-Credentialbox $SeqVariables.Split(';')[0] $SeqVariables.Split(';')[1] }
    }
'@

}


function Set-TabColor($ColorSelected, $ColorUnselected) {

  # Set the Tab Color

  $DataGridTabControl.SelectedTab.Tag.ColorSelected = $ColorSelected
  $DataGridTabControl.SelectedTab.Tag.ColorUnSelected = $ColorUnselected
  $DataGridTabControl.Refresh()

}


function Set-TabStyle {

  # Set the style of the Tabs

  if ($FormSettingsTabsLookCheckedRadioButton[0].Checked -eq $True) {
    # Set to normal (no colors)
    $DataGridTabControl.DrawMode = "Normal"
    $Script:TabLook = "0"
  }

  if ($FormSettingsTabsLookCheckedRadioButton[1].Checked -eq $True) {
    # Set to colors full
    $DataGridTabControl.DrawMode = "Normal"
    $DataGridTabControl.DrawMode = "OwnerDrawFixed"
    $DataGridTabControl.Remove_DrawItem($DataGridTabControl_DrawItemHandlerColorsFull)
    $DataGridTabControl.Remove_DrawItem($DataGridTabControl_DrawItemHandlerColorsLine)
    $DataGridTabControl.Add_DrawItem($DataGridTabControl_DrawItemHandlerColorsFull)
    $Script:TabLook = "1"
  }

  if ($FormSettingsTabsLookCheckedRadioButton[2].Checked -eq $True) {
    # Set to colors lines
    $DataGridTabControl.DrawMode = "Normal"
    $DataGridTabControl.DrawMode = "OwnerDrawFixed"
    $DataGridTabControl.Remove_DrawItem($DataGridTabControl_DrawItemHandlerColorsFull)
    $DataGridTabControl.Remove_DrawItem($DataGridTabControl_DrawItemHandlerColorsLine)
    $DataGridTabControl.Add_DrawItem($DataGridTabControl_DrawItemHandlerColorsLine)
    $Script:TabLook = "2"
  }

}


function Set-Timer {

  # Start the timer used to get a responsive GUI and get the state of the runspaces at every interval
 
  $Timer.Stop()
  $Form.Refresh()
  $Timer.Interval = $TimerInterval
  $Timer.Start()
  $Form.Refresh()

}


function Set-UnAssignSequenceToObjects {

  # Un-assign the selected objects from a group

  $GroupName = $OutputDataGrid.Rows[$OutputDataGrid.SelectedCells[0].RowIndex].Cells[0].Tag.GroupID  # Determine the group Name
  $SeqId = $OutputDataGrid.Rows[$OutputDataGrid.SelectedCells[0].RowIndex].Cells[8].Value  # and the respective sequence ID

  $DataGridViewCellStyleRegular = New-Object System.Windows.Forms.DataGridViewCellStyle
  $DataGridViewCellStyleRegular.Alignment = 16
  $DataGridViewCellStyleRegular.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [Drawing.FontStyle]::"Regular", 3, 0)
  $DataGridViewCellStyleRegular.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
  $DataGridViewCellStyleRegular.SelectionBackColor = [System.Drawing.Color]::FromArgb(255, 51, 153, 255)
  $DataGridViewCellStyleRegular.SelectionForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 255)
 
  foreach ($SelectedItem in $OutputDataGrid.SelectedCells) {
    # Reset the objects values to non-grouped objects values
    $OutputDataGrid.Rows[$SelectedItem.RowIndex].Cells[7].ReadOnly = $False
    $OutputDataGrid.Rows[$SelectedItem.RowIndex].Cells[0].Tag.GroupID = "0"
    $OutputDataGrid.Rows[$SelectedItem.RowIndex].Cells[0].Style = $DataGridViewCellStyleRegular
    $OutputDataGrid.Rows[$SelectedItem.RowIndex].Cells[9].Value = "-"
    if ($Sequences[$($OutputDataGrid.Rows[$SelectedItem.RowIndex].Cells[8].Value)].SequenceScheduler -ne 0) { 
      $OutputDataGrid.Rows[$SelectedItem.RowIndex].Cells[3].Value = "Pending" 
      $Script:SelectionChanged = $True
    }
  }

  $SequencesTreeView.SelectedNode = $SequencesTreeView.Nodes[0]  # Deselect the current selected sequence

  Set-RecreateGroups

  Get-CountCheckboxes

}


function Set-UserSettings {

  # Save the user's settings in the registry

  for ($i = 0; $i -le 3; $i++) {
    $ColorHex = "#FF{0:X2}{1:X2}{2:X2}" -f $FormSettingsColorsButton[$i].BackColor.R, $FormSettingsColorsButton[$i].BackColor.G, $FormSettingsColorsButton[$i].BackColor.B
    $Script:Colors.Set_Item($FormSettingsColorsButton[$i].Name, $ColorHex)
    $RegColorName = "Color_" + $FormSettingsColorsButton[$i].Name
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name $RegColorName -Value $ColorHex
  }

  for ($i = 0; $i -le 4; $i++) {
    Set-Variable -Name $FormSettingsPathsVariable[$i] -Value $FormSettingsPathsText[$i].Text -Scope Script -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name $FormSettingsPathsVariable[$i] -Value $FormSettingsPathsText[$i].Text
  }

  $FormSettingsMailText[2].Text = $FormSettingsMailText[2].Text -replace ";", ","  # Replace ";" with "," for multiple recipients
  for ($i = 0; $i -le 3; $i++) {
    Set-Variable -Name $FormSettingsMailVariable[$i] -Value $FormSettingsMailText[$i].Text -Scope Script -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name $FormSettingsMailVariable[$i] -Value $FormSettingsMailText[$i].Text
  }  

  $ColorHex = "#FF{0:X2}{1:X2}{2:X2}" -f $FormSettingsColorsGUIBackButton.BackColor.R, $FormSettingsColorsGUIBackButton.BackColor.G, $FormSettingsColorsGUIBackButton.BackColor.B
  Set-Variable -Name ColorBackground -Value $ColorHex -Scope Script -Force
  Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name ColorBackground -Value $ColorHex
  
  $ColorHex = "#FF{0:X2}{1:X2}{2:X2}" -f $FormSettingsColorsGUISeqButton.BackColor.R, $FormSettingsColorsGUISeqButton.BackColor.G, $FormSettingsColorsGUISeqButton.BackColor.B
  Set-Variable -Name ColorSequences -Value $ColorHex -Scope Script -Force
  Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name ColorSequences -Value $ColorHex

  $ColorHex = "#FF{0:X2}{1:X2}{2:X2}" -f $FormSettingsColorsGUISeqRunButton.BackColor.R, $FormSettingsColorsGUISeqRunButton.BackColor.G, $FormSettingsColorsGUISeqRunButton.BackColor.B
  Set-Variable -Name ColorSequencesRunning -Value $ColorHex -Scope Script -Force
  Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name ColorSequencesRunning -Value $ColorHex

  $Form.BackColor = $ColorBackground
  $SequencesTreeView.BackColor = $ColorSequences
  $SequenceTasksPanel.BackColor = $ColorSequences

  if ($FormSettingsSequenceShowSearchRadioButton.Checked) { 
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name ShowSearchBox -Value "True" -Force 
    Set-Variable -Name ShowSearchBox -Value "True" -Scope Script -Force
  }
  else { 
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name ShowSearchBox -Value "False" -Force
    Set-Variable -Name ShowSearchBox -Value "False" -Scope Script -Force
  }

  if ($FormSettingsSplashScreenCheckBox.Checked) { Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name NoSplashScreen -Value "False" -Force }
  else { Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name NoSplashScreen -Value "True" -Force }

  if ($FormSettingsDebugScreenCheckBox.Checked) { Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name DebugMode -Value 5 -Force }
  else { Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name DebugMode -Value 0 -Force }

  if ($FormSettingsSequenceExpandedRadioButton.Checked) { 
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name SequenceListExpanded -Value "True" -Force 
    $SequencesTreeView.ExpandAll()
  }
  else { 
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name SequenceListExpanded -Value "False" -Force 
    $SequencesTreeView.CollapseAll()
  }

  if ($FormSettingsRowHeaderGroupBoxVisibleRadioButton.Checked) { 
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name RowHeaderVisible -Value "True" -Force 
  }
  else { 
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name RowHeaderVisible -Value "False" -Force 
  }
  $OutputDataGrid.RowHeadersVisible = $FormSettingsRowHeaderGroupBoxVisibleRadioButton.Checked

  if ($FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked) { 
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name CheckBoxesKeepState -Value "True" -Force 
  }
  else { 
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name CheckBoxesKeepState -Value "False" -Force 
  }

  if ($FormSettingsLogCheckBox.Checked) { 
    Set-Variable -Name LogFileEnabled -Value "True" -Scope Script -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name LogFileEnabled -Value "True" -Force 
  }
  else { 
    Set-Variable -Name LogFileEnabled -Value "False" -Scope Script -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name LogFileEnabled -Value "False" -Force
  }

  if ($FormSettingsGroupsWarningCheckedRadioButton.Checked) { 
    Set-Variable -Name GrpCheckedOnWarning -Value "True" -Scope Script -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name GrpCheckedOnWarning -Value "True" -Force 
  }
  else { 
    Set-Variable -Name GrpCheckedOnWarning -Value "False" -Scope Script -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name GrpCheckedOnWarning -Value "False" -Force 
  }

  if ($FormSettingsGroupsThreadsVisibleRadioButton.Checked) { 
    Set-Variable -Name GrpShowThreads -Value "True" -Scope Script -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name GrpShowThreads -Value "True" -Force
  }
  else { 
    Set-Variable -Name GrpShowThreads -Value "False" -Scope Script -Force
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name GrpShowThreads -Value "False" -Force 
  }
  Show-GroupThreads

  Set-SearchBox
  Set-TabStyle
  Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name TabLook -Value $TabLook -Force

}


function Set-View_ColumnsSizeAuto {

  # Set the column size adjustment to fixed/auto
  
  $OutputDataGrid.Columns[0].Width = 100
  $OutputDataGrid.Columns[0].AutoSizeMode = 'Fill'
  $OutputDataGrid.Columns[0].FillWeight = 50
  $OutputDataGrid.Columns[2].Width = 150
  $OutputDataGrid.Columns[2].AutoSizeMode = 'Fill'
  $OutputDataGrid.Columns[2].FillWeight = 150
  $OutputDataGrid.Columns[3].Width = 100
  $OutputDataGrid.Columns[3].AutoSizeMode = 'None'
}


function Set-View_ColumnsSizeManual {

  # Set the column size adjustment to manual
  
  for ($i = 0; $i -le 3; $i++) { 
    $OutputDataGrid.Columns[$i].AutoSizeMode = 'None'
  }

}


function Set-View_Wrap {

  # Set/unset the cells content to wrap

  if (($MenuMain.Items.DropDown.items | where { $_.Text -eq "Wrap Text" }).Checked) {
    # Wrap mode set
    ($MenuMain.Items.DropDown.items | where { $_.Text -eq "Wrap Text" }).Checked = $False  # Deactivate the wrap mode
    $OutputDataGrid.RowsDefaultCellStyle.WrapMode = 'False'
    $OutputDataGrid.Columns[2].DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
    $OutputDataGrid.AutoSizeRowsMode = 'AllCellsExceptHeaders'
    $OutputDataGrid.Refresh()
    $OutputDataGrid.AutoSizeRowsMode = 'None'
  }
  else {
    # Activate the wrap mode
    ($MenuMain.Items.DropDown.items | where { $_.Text -eq "Wrap Text" }).Checked = $True
    $OutputDataGrid.RowsDefaultCellStyle.WrapMode = 'True'
    $OutputDataGrid.Columns[2].DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::TopLeft
    $OutputDataGrid.AutoSizeRowsMode = 'AllCellsExceptHeaders'
    $OutputDataGrid.Refresh()
  }

}


function Show-GroupThreads {

  # Display or hide the value of the max threads in the column Groups, depending on the user's settings

  foreach ($row in $OutputDataGrid.Rows) {
    if (($Row.Index -eq $OutputDataGrid.RowCount - 1) -or ($row.Cells[9].Value -eq "-")) { continue }
    $row.Cells[9].Value = $row.Cells[0].Tag.GroupID  # Get the name of the Group, in Cells[0].Tag.GroupID
    if ($GrpShowThreads -eq "True") { $row.Cells[9].Value += " ($($Sequences[$row.Cells[8].Value].MaxThreads))" }  # Show the threads values
  }

}


function Show-PickColor($Color) {

  # Show a color dialog and return the HEX value of the color chosen 

  $ColorDialog = New-Object System.Windows.Forms.ColorDialog
  $ColorDialog.AllowFullOpen = $true
  $ColorPicked = $ColorDialog.ShowDialog()
  
  if ($ColorPicked -eq "Cancel") {
    return "Cancel"
  }
  else {
    $HexColor = "#FF{0:X2}{1:X2}{2:X2}" -f $ColorDialog.Color.R, $ColorDialog.Color.G, $ColorDialog.Color.B
    $HexColor
  }

}


function Show-SequenceSteps {

  # Display the sequence steps of a sequence ID

  $SeqId = $OutputDataGrid.Rows[$OutputDataGrid.SelectedCells[0].RowIndex].Cells[8].Value  # Get the sequence ID based on the 1st object selected

  #Clear the sequence task panel and recreate all labels, checkboxes and variables based on the values of $ScriptBlockComment, $SequenceLabel, ScriptBlockCheckboxes and $ScriptBlockVariable of the Sequence ID
  $SequenceTasksPanel.Controls.Clear()

  $SequencePanelTitleLabel = New-Object System.Windows.Forms.Label
  $SequencePanelTitleLabel.Text = $($Sequences[$SeqId].SequenceLabel)
  $SequencePanelTitleLabel.Location = New-Object System.Drawing.Size(10, 10)
  $SequencePanelTitleLabel.AutoSize = $True
  $SequencePanelTitleLabel.MaximumSize = New-Object System.Drawing.Size(500, 15)
  $SequencePanelTitleLabel.Font = New-Object Drawing.Font("Tahoma", 9, [Drawing.FontStyle]::"Underline")
  $SequencePanelTitleLabel.ForeColor = [Drawing.Color]::"DarkViolet"
  $SequenceTasksPanel.Controls.Add($SequencePanelTitleLabel)

  for ($SeqPosition = 0; $SeqPosition -lt $Sequences[$SeqId].ScriptBlockComment.Count; $SeqPosition++) {
    $SequencePanelTempCheckbox = $Sequences[$SeqId].ScriptBlockCheckboxes[$SeqPosition]
    $Sequences[$SeqId].ScriptBlockCheckboxes[$SeqPosition] = New-Object System.Windows.Forms.CheckBox
    $Sequences[$SeqId].ScriptBlockCheckboxes[$SeqPosition].ForeColor = [Drawing.Color]::"DarkViolet"
    $Sequences[$SeqId].ScriptBlockCheckboxes[$SeqPosition].Location = New-Object System.Drawing.Size($SequencePanelTempCheckbox.Left, $SequencePanelTempCheckbox.Top)
    $Sequences[$SeqId].ScriptBlockCheckboxes[$SeqPosition].Checked = $SequencePanelTempCheckbox.Checked
    $Sequences[$SeqId].ScriptBlockCheckboxes[$SeqPosition].AutoSize = $True
    $Sequences[$SeqId].ScriptBlockCheckboxes[$SeqPosition].MaximumSize = New-Object System.Drawing.Size(500, 15)
    $Sequences[$SeqId].ScriptBlockCheckboxes[$SeqPosition].Font = New-Object Drawing.Font("Microsoft Sans Serif", 8, [Drawing.FontStyle]::"Regular")
    $Sequences[$SeqId].ScriptBlockCheckboxes[$SeqPosition].Text = $SequencePanelTempCheckbox.Text
    $SequenceTasksPanel.Controls.Add($Sequences[$SeqId].ScriptBlockCheckboxes[$SeqPosition])
    $SequencePanelTempLabel = New-Object System.Windows.Forms.Label
    $SequencePanelTempLabel.Text = "  $($Sequences[$SeqId].ScriptBlockComment[$SeqPosition])`n`n"
    $Pos_Y = $(43 * $SeqPosition + 62)
    $SequencePanelTempLabel.Location = New-Object System.Drawing.Size(15, $Pos_Y)
    $SequencePanelTempLabel.AutoSize = $True
    $SequencePanelTempLabel.MaximumSize = New-Object System.Drawing.Size(500, 15)
    $SequencePanelTempLabel.Font = New-Object Drawing.Font("Microsoft Sans Serif", 8, [Drawing.FontStyle]::"Italic")
    $SequencePanelTempLabel.ForeColor = [Drawing.Color]::"DarkViolet"
    $SequenceTasksPanel.Controls.Add($SequencePanelTempLabel)
  }

  $VarPos = 0
  for ($i = 1; $i -le $nbVariableTypes; $i++) {
    # Parse all types of variables known
    $Sequences[$SeqId].ScriptBlockVariable[$i].keys | foreach { # Check if variables have been defined for this type
      $VarName = $_  # Get the variable name
      $VarValue = $Sequences[$SeqId].ScriptBlockVariable[$i].Item($_)  # Get the variable value
      $SequencePanelTempVariable = New-Object System.Windows.Forms.Label
      $SequencePanelTempVariable.Text = "$VarName : $VarValue" -replace "`n", "|"  # Write everything on one line only
      $Pos_Y = 44 * $($Sequences[$SeqId].ScriptBlockComment.Count) + 16 * $VarPos + 48
      $SequencePanelTempVariable.Location = New-Object System.Drawing.Size(15, $Pos_Y)
      $SequencePanelTempVariable.AutoSize = $True
      $SequencePanelTempVariable.MaximumSize = New-Object System.Drawing.Size(500, 15)
      $SequencePanelTempVariable.Font = New-Object Drawing.Font("Microsoft Sans Serif", 8, [Drawing.FontStyle]::"Italic")
      $SequencePanelTempVariable.ForeColor = [Drawing.Color]::"Blue"
      $SequenceTasksPanel.Controls.Add($SequencePanelTempVariable)
      $VarPos++
    }
  }

}


function Start-Sequence {

  # Start the sequence

  if (Set-AssignSequenceToFreeObjects -eq "err") { return }  # Nothing to assign or run, exit
  Set-RunspaceToObjects  # Define the Runspaces for the sequences to run

  $Script:SelectionChanged = $False

  $OutputDataGrid.Focus()
  foreach ($Item in $SequencesToParse) {
    # Loop into the ID of the sequences to run
    $CheckedObjects = @($OutputDataGridSequence[$Item].Cells | where { ($_.ColumnIndex -eq 7) -and ($_.Value -eq $True) }).Count
    if ( ($CheckedObjects -gt $Sequences[$Item].MaxCheckedObjects) -and ($Sequences[$Item].MaxCheckedObjects -ne 0) ) {
      # Too much objects selected
      [System.Windows.Forms.MessageBox]::Show("Too much objects selected for '$($Sequences[$Item].SequenceLabel)'`n`nMaximum allowed: $($Sequences[$Item].MaxCheckedObjects), Selected: $CheckedObjects", "$($Sequences[$Item].SequenceLabel)", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Stop) 
      $OutputDataGridSequence[$Item].Cells | where { $_.ColumnIndex -eq 7 } | foreach { $_.Value = $False }  # Uncheck all the objects of the sequence to skip it
      $OutputDataGrid.RefreshEdit()
    }

    $OutputDataGridSequence_DataGridViewCellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $OutputDataGridSequence_DataGridViewCellStyle.Alignment = 16
    $OutputDataGridSequence_DataGridViewCellStyle.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, 0, 3, 0)
    $OutputDataGridSequence_DataGridViewCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0)
    $OutputDataGridSequence_DataGridViewCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(255, 51, 153, 255)
    $OutputDataGridSequence_DataGridViewCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 255)

    foreach ($Row in $OutputDataGridSequence[$Item]) {
      # Loop into each row of each sequence ID
      if (($Row.Cells[7].Value) -and ($row.Cells[4].Value -le 0) -and ($row.Cells[4].Value -ne -5)) {
        # The object is checked, not running and not cancelling
        $Row.Cells[0].ReadOnly = $True  # Make the object name non-editable during the sequence run
        $Row.Cells[7].ReadOnly = $True  # Make the checkbox non-clickable during the sequence run
        $Row.DefaultCellStyle.BackColor = $ColorSequencesRunning
        $Row.Cells[0].Tag.StepProtocol = @()  # Reset the protocol
        $Row.Cells[2].Style = $OutputDataGridSequence_DataGridViewCellStyle
        $Row.Cells[3].Value = "Pending"
        $Row.Cells[4].Value = 0  # Reset the Step ID to 0
        $Row.Cells[0].Tag.PreviousStateComment = ""  # Reset the previous state comment
        $Row.Cells[0].Tag.SharedVariable = $Null  # Reset the shared variable between the steps#>
      }
    }
  }

  $ActionButton.Enabled = $False  # Disable the Start button
  ($MenuToolStrip.Items | where { $_.ToolTipText -eq "Clear the Grid" }).Enabled = $False  # Disable the "Clear the Grid" icon

  if (!($SequenceRunning)) {
    # No other sequence is currently running
    $Script:SequenceRunning = $True
    $Script:ObjectsDone = 0
    $Timer.Enabled = $True  # Enable the Timer
    ($MenuMain.Items | where { $_.Text -eq "Cancel" }).Enabled = $True  # Enable and disable some menu items
    ($MenuToolStrip.Items | where { $_.ToolTipText -eq "Cancel All" }).Enabled = $True
    ($MenuMain.Items.DropDown.Items | where { $_.Text -eq "Reset All Objects" }).Enabled = $False
    ($MenuMain.Items.DropDown.Items | where { $_.Text -eq "Clear Grid" }).Enabled = $False
    Set-Timer
  }

  if ($TimerSet -eq $False) {
    # Timer not activated
    $Timer.Add_Tick($GetData)  # Start $GetDate at every Tick of the Timer
    $Script:TimerSet = $True 
  }

} 


function Write-DebugReceiveOutput($ReceiveOutput) {

  for ($i = 0; $i -lt $ReceiveOutput.Count; $i++) {
    switch ($i) {
      0 { Write-Host "Value 1 - Status: $($ReceiveOutput[$i])" }
      1 { Write-Host "Value 2 - Result state: $($ReceiveOutput[$i])" }
      2 { Write-Host "Value 3 - Color: $($ReceiveOutput[$i])" }
      3 { Write-Host "Value 4 - Shared value: $($ReceiveOutput[$i])" }
      { $_ -gt 3 } { Write-Host "Value $($i+1) - Error: $($ReceiveOutput[$i])" }
    } 
  }

}


function Write-Log {

  # Write a log 

  if ($LogFileEnabled -eq $False) { return }

  $ToLog = ""
  foreach ($Item in $SequencesToParse) {
    $ToLog += "$($Sequences[$Item].SequenceLabel)  -  $((Get-Date).ToShortDateString())`r`n"
    foreach ($Row in $OutputDataGridSequence[$Item]) {
      if ($Row.Cells[0].Tag.StepProtocol -ne $Null) {
        $ToLog += "$(($Row.Cells[0].Tag.StepProtocol).Trim() -join " ; ") `r`n"
      }
    }
    $ToLog += "`r`n`r`n"
  }
  Add-Content -Value $ToLog -Path $LogFilePath

}


#-----------------

$HydraVersion = "5.55"

Set-StrictMode -Version Latest

Set-StartupSettings $MyInvocation

[Reflection.Assembly]::Loadwithpartialname("System.Windows.Forms") | Out-Null
[Reflection.Assembly]::Loadwithpartialname("System.Drawing") | Out-Null
[Reflection.Assembly]::Loadwithpartialname("PresentationFramework") | Out-Null
 
[System.Windows.Forms.Application]::EnableVisualStyles() | Out-Null

. $HydraGUIPath\Hydra5_Res.ps1
. $HydraGUIPath\Hydra5_Form.ps1
. $HydraGUIPath\Dialogs.ps1

. Load-Logo
if ($NoSplashScreen -ne "True") { [void]$FormSplashScreen.ShowDialog() }

. Add-Form

$Form.ShowDialog() | Out-Null

[void] $Form.Dispose()
$Timer.Dispose()
