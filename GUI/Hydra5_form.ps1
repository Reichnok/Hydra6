function Add-Form {

  function Add-GridColumn($GridTab, $AutoSizeMode, $Name, $MinWidth, $Width, $Visible, $ReadOnly) {

    # Create a new Column
    $DataGrid_TextBoxColumn=New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $DataGrid_TextBoxColumn.AutoSizeMode=$AutoSizeMode
    $DataGrid_TextBoxColumn.HeaderText=$Name
    $DataGrid_TextBoxColumn.MinimumWidth=$MinWidth
    $DataGrid_TextBoxColumn.Name=$Name
    $DataGrid_TextBoxColumn.Width=$Width
    $DataGrid_TextBoxColumn.Visible=$Visible
    $DataGrid_TextBoxColumn.ReadOnly=$ReadOnly
    $OutputDataGridTab[$GridTab].Columns.Add($DataGrid_TextBoxColumn) | Out-Null
  }

  function Add-ToolStripButton($Text, $Icon, $Script, $MenuStrip, $Size) {

    # Create a new icon in the Icon Bar
    $MenuToolStripButton=New-Object System.Windows.Forms.ToolStripButton
    $MenuToolStripButton.AutoSize=$False
    $MenuToolStripButton.Height=$Size
    $MenuToolStripButton.Width=$Size
    $MenuToolStripButton.ToolTipText=$Text
    $MenuToolStripButton.Name=$Text
    $MenuToolStripButton.Image=$Icon
    $MenuToolStripButton.Add_Click( $Script  )
    [void]$MenuStrip.Items.Add($MenuToolStripButton)
  }

  function Add-ToolStripSeparator {

    # Create a new separator in the Icon Bar
    $MenuToolStripSeparator=New-Object System.Windows.Forms.ToolStripSeparator
    $MenuToolStripSeparator.Margin=7
    $MenuToolStripSeparator.Visible=$True
    [void]$MenuToolStrip.Items.Add($MenuToolStripSeparator)
  }

  function Add-MenuItem($MainMenuName, $SubMenus) {

    # Add a new entry in the Main Menu: use a hash collection as parameter
    $MainMenuItem=New-Object System.Windows.Forms.ToolStripMenuItem
    $MainMenuItem.Text=$MainMenuName
    $MainMenuItem.Name=$MainMenuName
    [void]$MenuMain.Items.Add($MainMenuItem)

    foreach($item in $SubMenus.GetEnumerator()) {  # Parse the hashes
      if ($item.Value-eq "-") {
        $MenuSubItem=New-Object System.Windows.Forms.ToolStripSeparator
      }
      else {
        $MenuSubItem=New-Object System.Windows.Forms.ToolStripMenuItem
        $MenuSubItem.Text=$item.Name
        $MenuSubItem.Add_Click($item.Value)
      }
      [void]$MainMenuItem.DropDownItems.Add($MenuSubItem)
    }
  }

  function Add-ContextMenuStripItem($ContextMenuStrip, $Text, $Script) {

    # Add entries in a context menu
    $ContextMenuStripItem=New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ContextMenuStripItem.Name=$Text
	$ContextMenuStripItem.Text=$Text
    $ContextMenuStripItem.Add_Click( $Script )  
    [void]$ContextMenuStrip.Items.Add($ContextMenuStripItem) 
  }

  function Add-ContextSubMenuStripItem($ContextMenuStrip, $Text, $Script, $Enabled=$True, $Color=$Null) {

    # Add entries in a submenu of a context menu
    $ContextSubMenuStripItem=New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ContextSubMenuStripItem.Name=$Text
	$ContextSubMenuStripItem.Text=$Text
    $ContextSubMenuStripItem.Enabled=$Enabled
    if ($Color -ne $Null) { $ContextSubMenuStripItem.BackColor=$Color }
    $ContextSubMenuStripItem.Add_Click( $Script )  
    [void]$ContextMenuStrip.DropDownItems.Add($ContextSubMenuStripItem) 
  }

  function Add-ContextStripSeparator ($ContextMenuStrip){

    # Add a separator in a context menu
    $StripSeparator=New-Object 'System.Windows.Forms.ToolStripSeparator'
    [void]$ContextMenuStrip.Items.Add($StripSeparator)
  }

  function Add-ContextSubStripSeparator ($ContextMenuStrip){

    # Add a separator in a submenu of a context menu
    $StripSeparator=New-Object 'System.Windows.Forms.ToolStripSeparator'
    [void]$ContextMenuStrip.DropDownItems.Add($StripSeparator)
  }

  function SequencesTreeView_Filter($Filter) {

    # Filter the sequences based on the search box criteria
    $SequencesTreeView.Nodes.Clear()
    $SequenceListRootNode=New-Object System.Windows.Forms.TreeNode
    $SequenceListRootNode.Text="Filtered"
    $SequenceListRootNode.Name="Filtered"
    [void]$SequencesTreeView.Nodes.Add($SequenceListRootNode)

    $i=0
    foreach($Sequence in $SequenceList) {  # Loop in the filtered sequences and create the entries
      if (($Sequence.SeqName -match $Filter) -and ($Sequence.SeqName -notmatch "-----")) {
      $SequenceListSubNode=New-Object -TypeName System.Windows.Forms.TreeNode
      $SequenceListSubNode.Text=$Sequence.SeqName  
      $SequenceListSubNode.Tag=$i                   
      [void]$SequenceListRootNode.Nodes.Add($SequenceListSubNode)
      }

      $i++
    }
    $SequencesTreeView.ExpandAll()
  }

  function SequencesTreeView_GetSequenceList($ReloadSequence) {

    # Load the Sequences names and paths from the Sequence List
    if ($ReloadSequence) {
      $Script:SequenceList=Import-Csv -Delimiter ";" -Path $SequencesListPath -Header SeqName, SeqPath  # Create the global variable SequenceList with names and paths
    }
    $i=0
    foreach($Sequence in $SequenceList) {  # Loop in the sequences and create the entries
      if ($Sequence.SeqName -like "*-----*") {  # Separator found, create a dummy entry
        $SequenceListRootNode=New-Object System.Windows.Forms.TreeNode
        $SequenceListRootNode.Text=$Sequence.SeqName -replace "-----"; "" | Out-Null
        $SequenceListRootNode.Name=$Sequence.SeqName -replace "-----"; "" | Out-Null
        if ($SequenceListExpanded -eq "True") { $SequenceListRootNode.ExpandAll() } else { $SequenceListRootNode.Collapse() }
        [void]$SequencesTreeView.Nodes.Add($SequenceListRootNode)
      }
      else {  # Create the new entry and set its position in the node's tag
        $SequenceListSubNode=New-Object -TypeName System.Windows.Forms.TreeNode
        $SequenceListSubNode.Text=$Sequence.SeqName  
        $SequenceListSubNode.Tag=$i                
        [void]$SequenceListRootNode.Nodes.Add($SequenceListSubNode)
      }
      $i++
    }
  }

  function Set-SubMenuIcons ($SubMenuLabel, $IconList){

    # Set the icons of the entries in the Main Menu: use a collection as parameter
    $SubMenu=($MenuMain.Items | where { $_.Text -eq $SubMenuLabel })
    for ($i=0; $i -lt $IconList.Count; $i++) {
      $SubMenu.DropDownItems[$i].Image=$IconList[$i]
    }
  }

  function Set-NewTab {

    # Create a new Tab, and the grid associated to
    $Script:TabPageIndex++
    $TabPage=New-Object System.Windows.Forms.TabPage  # Create the new TabPage and set its attributes
    $Tabpage.Location='4, 22'
    $TabPage.Name="tabpage$TabPageIndex"
    $TabPage.Padding='3, 3, 3, 3'
    $TabPage.Anchor='Top, Bottom, Left, Right'
    $TabPage.TabIndex=$TabPageIndex
    $TabPage.Text="  Objects List $TabPageIndex  "
    $TabPageOptions=New-Object -TypeName PSObject
    $TabPageOptions | Add-Member –MemberType NoteProperty –Name TabPageIndex –Value $TabPageIndex
    $TabPageOptions | Add-Member –MemberType NoteProperty –Name ColorSelected –Value "DarkGray"
    $TabPageOptions | Add-Member –MemberType NoteProperty –Name ColorUnSelected –Value "LightGray"
    $TabPage.Tag=$TabPageOptions
    $TabPage.UseVisualStyleBackColor=$True
    $DataGridTabControl.Controls.Add($TabPage)
    $TabPage.Add_Paint({  # Paint the header in the $ColorBackground instead of grey
      $Brush=New-Object Drawing.SolidBrush $ColorBackground
      $TabWidth=$DataGridTabControl.GetTabRect($DataGridTabControl.TabCount-1).Right
      $DataGridTabControl.CreateGraphics().FillRectangle($Brush, $TabWidth, 0, $DataGridTabControl.Width, 20)
    })

    $Script:OutputDataGridTab+=New-Object System.Windows.Forms.DataGridView  # Create a new DataGridView and define all its attributes
    # Value :  0: Objects ; 1: JobID ; 2: Results ; 3: Step ; 4: StepID ; 5: Color ; 6: FileSource ; 7: Checkbox ; 8: SequenceID ; 9: Group
    # Tag[0]:  GroupID ; PreviousStateComment ; StepProtocol ; SharedVariable
    $Script:OutputDataGridTab[$TabPageIndex].Anchor='Top, Bottom, Left, Right' 
    $OutputDataGrid_CheckBoxColumn=New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $OutputDataGrid_CheckBoxColumn.FlatStyle=1
    $OutputDataGrid_CheckBoxColumn.Name=""
    $OutputDataGrid_CheckBoxColumn.ReadOnly=$True
    $OutputDataGrid_CheckBoxColumn.Resizable=2
    $OutputDataGrid_CheckBoxColumn.Width=30
    $OutputDataGrid_CheckBoxColumn.DisplayIndex=0
    Add-GridColumn $TabPageIndex 4 "Objects" 100 100 $True $False
    Add-GridColumn $TabPageIndex 1 "JobId" 50 50 $False $True
    Add-GridColumn $TabPageIndex 16 "Task Result" 200 594 $True $True
    Add-GridColumn $TabPageIndex 1 "State" 100 100 $True $True
    Add-GridColumn $TabPageIndex 1 "StepID" 50 50 $False $True
    Add-GridColumn $TabPageIndex 1 "Color" 10 10 $False $True
    Add-GridColumn $TabPageIndex 1 "FileSource" 10 10 $False $True
    $Script:OutputDataGridTab[$TabPageIndex].Columns.Add($OutputDataGrid_CheckBoxColumn) | Out-Null
    Add-GridColumn $TabPageIndex 1 "SequenceID" 50 50 $False $True
    Add-GridColumn $TabPageIndex 1 "Group" 80 50 $False $True
    $OutputDataGrid_DataGridViewCellStyle=New-Object System.Windows.Forms.DataGridViewCellStyle
    $OutputDataGrid_DataGridViewCellStyle.Alignment=16
    $OutputDataGrid_DataGridViewCellStyle.BackColor=[System.Drawing.Color]::FromArgb(255,255,255,255)
    $OutputDataGrid_DataGridViewCellStyle.Font=New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,0,3,0)
    $OutputDataGrid_DataGridViewCellStyle.ForeColor=[System.Drawing.Color]::FromArgb(255,0,0,0)
    $OutputDataGrid_DataGridViewCellStyle.SelectionBackColor=[System.Drawing.Color]::FromArgb(255,51,153,255)
    $OutputDataGrid_DataGridViewCellStyle.SelectionForeColor=[System.Drawing.Color]::FromArgb(255,255,255,255)
    $OutputDataGrid_DataGridViewCellStyle.WrapMode=2
    $OutputDataGrid_DataGridViewHeaderStyle=New-Object System.Windows.Forms.DataGridViewCellStyle
    $OutputDataGrid_DataGridViewHeaderStyle.Alignment=16
    $OutputDataGrid_DataGridViewHeaderStyle.BackColor=[System.Drawing.Color]::FromArgb(255,192,192,192)
    $OutputDataGrid_DataGridViewHeaderStyle.Font=New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,0,3,0)
    $OutputDataGrid_DataGridViewHeaderStyle.ForeColor=[System.Drawing.Color]::FromArgb(255,0,0,0)
    $OutputDataGrid_DataGridViewHeaderStyle.WrapMode=1
    $Script:OutputDataGridTab[$TabPageIndex].DefaultCellStyle=$OutputDataGrid_DataGridViewCellStyle  # Define the attributes of the new TabPage
    $Script:OutputDataGridTab[$TabPageIndex].ColumnHeadersDefaultCellStyle=$OutputDataGrid_DataGridViewHeaderStyle
    $Script:OutputDataGridTab[$TabPageIndex].Location="0,0"
    $Script:OutputDataGridTab[$TabPageIndex].Name="OutputDataGrid"
    $Script:OutputDataGridTab[$TabPageIndex].RowHeadersVisible=($RowHeaderVisible -eq "True")
    $Script:OutputDataGridTab[$TabPageIndex].TabIndex=$TabPageIndex
    $Script:OutputDataGridTab[$TabPageIndex].StandardTab=$False
    $Script:OutputDataGridTab[$TabPageIndex].AllowUserToDeleteRows=$True
    $Script:OutputDataGridTab[$TabPageIndex].Width=$([int]($DataGridTabControl.Width) -10)
    $Script:OutputDataGridTab[$TabPageIndex].Height=$([int]($DataGridTabControl.Height) -24)
    $Script:OutputDataGridTab[$TabPageIndex].Tag=$TabPageOptions
    $Script:OutputDataGridTab[$TabPageIndex].BorderStyle="None"
    $TabPage.Controls.Add($OutputDataGridTab[$TabPageIndex])  # Add the new TabPage
    $DataGridTabControl.ResumeLayout()
    $Form.Refresh()

  }

  # Main Form and components: Use the saved parameters for the sizes and positions of the different elements

  $Form=New-Object System.Windows.Forms.Form
  $Form.SuspendLayout()
  $Form.ClientSize="$PosFormW, $PosFormH"
  $Form.MinimumSize="1150,750"
  $Form.Name="Form"
  if ([Environment]::Is64BitProcess) { $env="x64" } else { $env="x32" }
  $Form.Text="Hydra $HydraVersion  -  " +[Environment]::UserName + "    ($env)"
  $Form.StartPosition="Manual"
  $Form.Top=$PosFormY
  $Form.Left=$PosFormX
  $Form.MinimizeBox=$True
  $Form.MaximizeBox=$True
  $Form.WindowState="Normal"
  $Form.SizeGripStyle="Hide"
  $Form.BackColor=$ColorBackground
  $Form.Add_FormClosed( { Set-CloseForm } )
  $Form.Icon=$HydraIcon

  $FormBorderSize=$($Form.Width - $Form.ClientSize.Width) /2  # Correction for the window size
  $FormHeaderSize=$($Form.Height - $Form.ClientSize.height - 2* $FormBorderSize)  # Correction for the window size

  #Split Containers
  $SplitContainer1=New-Object 'System.Windows.Forms.SplitContainer'
  $SplitContainer2=New-Object 'System.Windows.Forms.SplitContainer'
  $SplitContainer1.Dock='Fill'
  $SplitContainer1.Location='0, 24'
  $SplitContainer1.Name='SplitContainer1'
  $SplitContainer1.Panel1.AutoScroll=$True
  $SplitContainer1.Panel1.AutoScrollMinSize='215, 0'
  [void]$SplitContainer1.Panel1.Controls.Add($SplitContainer2)
  $SplitContainer1.Panel1MinSize=230
  $SplitContainer1.Width=$PosSplit1W
  $SplitContainer1.Height=$PosSplit1H
  $SplitContainer1.SplitterDistance=$PosSplit1D
  $SplitContainer1.FixedPanel='Panel1'
  $Form.Controls.Add($SplitContainer1)

  $SplitContainer2.Dock='Fill'
  $SplitContainer2.Location='0, 0'
  $SplitContainer2.Name='SplitContainer2'
  $SplitContainer2.Orientation='Horizontal'
  $SplitContainer2.Panel1MinSize=75
  $SplitContainer2.Panel2MinSize=75
  $SplitContainer2.Width=$PosSplit2W
  $SplitContainer2.Height=$PosSplit2H
  $SplitContainer2.SplitterDistance=$PosSplit2D

  # Sequences panels
  if ($ShowSearchBox -eq "True") { $SequencesTreeViewTopPosition=20 } else { $SequencesTreeViewTopPosition=0 }
  $SequencesTreeView=New-Object System.Windows.Forms.TreeView
  $SequencesTreeView.Top=[int]$SeqTreeTop + $SequencesTreeViewTopPosition 
  $SequencesTreeView.Left=$SeqTreeLeft
  $SequencesTreeView.Name="SequencesTreeView"
  $SequencesTreeView.Width=$SeqTreeWidth
  $SequencesTreeView.Height=$SeqTreeHeight - $SequencesTreeViewTopPosition
  $SequencesTreeView.HideSelection=$False
  $SequencesTreeView.Anchor='Top,Left,Right,Bottom'
  $SequencesTreeView.TabIndex=0
  $SequencesTreeView.BackColor=$ColorSequences
  [void]$SplitContainer2.Panel1.Controls.Add($SequencesTreeView)

  $SearchTreeTextBox=New-Object System.Windows.Forms.TextBox
  $SearchTreeTextBox.Top=$([int]$SeqTreeTop + 1)
  $SearchTreeTextBox.Left=$([int]$SeqTreeLeft+1)
  $SearchTreeTextBox.Width=$([int]$SeqTreeWidth-2)
  $SearchTreeTextBox.Height=20
  $SearchTreeTextBox.Text="Search..."
  $SearchTreeTextBox.ForeColor="Gray"
  $SearchTreeTextBox.BorderStyle="Fixed3d"
  $SearchTreeTextBox.Visible=($ShowSearchBox -eq "True")
  [void]$SplitContainer2.Panel1.Controls.Add($SearchTreeTextBox)
  $SearchTreeTextBox.Add_TextChanged({
    try {
      $CurrentNode=$SequencesTreeView.SelectedNode.Text
    }
    catch {
      $CurrentNode=$Null
    }
    if (($SearchTreeTextBox.Text.Length -eq 0) -and ($SequencesTreeView.Nodes[0].Name -eq "Filtered")) { 
      $SequencesTreeView.Nodes.Clear()
      SequencesTreeView_GetSequenceList $False 
      if ($FormSettingsSequenceExpandedRadioButton.Checked) {  # Collapse or expand depending on the user's settings
        $SequencesTreeView.ExpandAll()
      }
      else { 
        $SequencesTreeView.CollapseAll()
      }
    }
    elseif (($SearchTreeTextBox.Text -ne "Search...") -and ($SearchTreeTextBox.Text.Length -ne 0)) { SequencesTreeView_Filter $SearchTreeTextBox.Text}
    try {
      foreach ($item in $SequencesTreeView.Nodes.Nodes) { 
        if ($item.Text -eq $CurrentNode) { 
          $SequencesTreeView.SelectedNode=$item
        }
      }
    }
    catch { }
  })

  $SearchTreeTextBox.Add_Enter({
    if ($SearchTreeTextBox.Text -eq "Search...") {
      $SearchTreeTextBox.Text=""
    }
  })

  $SearchTreeTextBox.Add_Leave({
    if ($SearchTreeTextBox.Text -eq "") {
      $SearchTreeTextBox.Text="Search..."
    }
  })

  SequencesTreeView_GetSequenceList $True  # Load the sequences

  $SequencesTreeView_AfterSelect={  # Actions to execute after a new sequence has been selected
    $Script:UseScheduler=$False
    $Script:SelectionChanged=$True  # Set the global variable SelectionChanged to True for further steps
    if ($SequencesTreeView.SelectedNode.Parent -ne $NULL) {  # Not a separator
      Get-Sequence $SequenceList[$SequencesTreeView.SelectedNode.Tag].SeqPath $SequencesTreeView.SelectedNode.Text  # Load the sequence steps
    }
    else {  # Separator
      $SequenceTasksPanel.Controls.Clear()
      $Script:SequenceLoaded=$False  # No Sequence loaded
      Set-ActionButtonState  # Enable or disable the different action components
    }
  }
  $SequencesTreeView.Add_AfterSelect( $SequencesTreeView_AfterSelect )

  $SequencesTreeView_MouseClickHandler=[System.Windows.Forms.MouseEventHandler]{  # Creation of the Sequence right click menu
    
    if ($_.Button -ne "Right") { return }  # Not a right click, exit
    if ($OutputDataGrid.Columns[9].Visible -eq $False) { $SequencesTreeView.ContextMenuStrip=$Null ; return }  # No Column 9 visible, no group set at all, exit
    
    $SequencesTreeViewContextMenu=New-Object 'System.Windows.Forms.ContextMenuStrip'  # Create the right click context menu
    $SequencesTreeView.ContextMenuStrip=$SequencesTreeViewContextMenu
    Add-ContextMenuStripItem $SequencesTreeViewContextMenu "Assign to ..." { }
    Add-ContextMenuStripItem $SequencesTreeViewContextMenu "Assign With Scheduler to ..." { }
    $SequencesTreeViewContextMenuAssign=$SequencesTreeViewContextMenu.Items | Where { $_.Text -eq "Assign to ..." }
    $SequencesTreeViewContextMenuAssignSched=$SequencesTreeViewContextMenu.Items | Where { $_.Text -eq "Assign With Scheduler to ..." }
    $GroupsInCurrentGrid=@(($OutputDataGrid.Rows.Cells | Where { ($_.ColumnIndex -eq 0) } | Select -ExpandProperty Tag) | Where { $_.GroupID -ne 0 } | Select -ExpandProperty GroupID -Unique | Sort-Object)  # Enumerate the groups in the grid
    foreach ($GroupUsedItem in $GroupsInCurrentGrid) {  # Create an entry for each group: one without and with scheduler option
      if ($GroupUsedItem -eq "0") { continue }  # Group 0 found: no group, exit
      $SBAssign=[scriptblock]::Create("Set-AssignSequenceToObjects $False $GroupUsedItem")
      Add-ContextSubMenuStripItem $SequencesTreeViewContextMenuAssign $GroupUsedItem $SBAssign ($GroupUsedItem -notin $GroupsRunning)
      $SBAssignSched=[scriptblock]::Create("Set-AssignSequenceToObjects $True $GroupUsedItem")
      Add-ContextSubMenuStripItem $SequencesTreeViewContextMenuAssignSched $GroupUsedItem $SBAssignSched ($GroupUsedItem -notin $GroupsRunning)
    }

  }
  $SequencesTreeView.Add_MouseClick( $SequencesTreeView_MouseClickHandler )

  $SequencesLabel=New-Object System.Windows.Forms.Label
  $SequencesLabel.BackColor=[System.Drawing.Color]::FromArgb(255,192,192,192)
  $SequencesLabel.Location="10,15"
  $SequencesLabel.Name="SequencesLabel"
  $SequencesLabel.Width=$SeqTreeWidth
  $SequencesLabel.Height=20
  $SequencesLabel.BorderStyle="Fixed3d"
  $SequencesLabel.Text="Sequences"
  $SequencesLabel.TextAlign="MiddleLeft"
  $Font=New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
  $SequencesLabel.Font=$Font
  $SequencesLabel.Anchor='Top,Left,Right'
  [void]$SplitContainer2.Panel1.Controls.Add($SequencesLabel)

  $SequenceTasksLabel=New-Object System.Windows.Forms.Label
  $SequenceTasksLabel.BackColor=[System.Drawing.Color]::FromArgb(255,192,192,192)
  $SequenceTasksLabel.Location="10,5"
  $SequenceTasksLabel.Name="SequenceTasksLabel"
  $SequenceTasksLabel.Width=$SeqTreeWidth
  $SequenceTasksLabel.Height=20
  $SequenceTasksLabel.BorderStyle="Fixed3d"
  $SequenceTasksLabel.TextAlign="MiddleLeft"
  $SequenceTasksLabel.Text="Sequence Steps"
  $Font=New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
  $SequenceTasksLabel.Font=$Font
  $SequenceTasksLabel.Anchor='Top,Left,Right'
  [void]$SplitContainer2.Panel2.Controls.Add($SequenceTasksLabel)

  $SequenceTasksPanel=New-Object System.Windows.Forms.Panel
  $SequenceTasksPanel.Top=$SeqPanelTop 
  $SequenceTasksPanel.Left=$SeqPanelLeft
  $SequenceTasksPanel.Width=$SeqPanelWidth
  $SequenceTasksPanel.Height=$SeqPanelHeight
  $SequenceTasksPanel.Name="SequenceTasksPanel"
  $SequenceTasksPanel.Anchor='Bottom,Top,Left,Right'
  $SequenceTasksPanel.AutoScroll=$True
  $SequenceTasksPanel.TabIndex=2
  $SequenceTasksPanel.BackColor=$ColorSequences
  [void]$SplitContainer2.Panel2.Controls.Add($SequenceTasksPanel)

  # Tab Control
  $DataGridTabControl=New-Object System.Windows.Forms.TabControl
  $DataGridTabControl.Location=New-Object System.Drawing.Size(10, 15)
  $DataGridTabControl.Width=$DataGridTabControlWidth
  $DataGridTabControl.Height=$DataGridTabControlHeight
  $DataGridTabControl.Name='TabControlMulti'
  $DataGridTabControl.Anchor='Top, Bottom, Left, Right'
  $DataGridTabControl.SelectedIndex=0
  $DataGridTabControl.TabIndex=0
  $DataGridTabControl.Appearance=0
  $DataGridTabControl.DrawMode="OwnerDrawFixed"
  $SplitContainer1.Panel2.Controls.Add($DataGridTabControl)

  # Create the 1st Tab and its grid and assign this Grid to the main variable $OutputDataGrid
  $OutputDataGridTab=,@()
  $Script:TabPageIndex=0
  Set-NewTab
  $Script:GridIndex=1
  $OutputDataGrid=$OutputDataGridTab[1]

  $DataGridTabControl_DrawItemHandlerColorsLine={  # Set the Tab Colors and Text
   
    $CurrentTabIndex=$DataGridTabControl.SelectedTab.TabIndex

    if ($($DataGridTabControl.SelectedIndex) -ne $($_.index)) { return }  # Skip useless repaints

    $Font=New-Object Drawing.Font("Microsoft Sans Serif", 8, [Drawing.FontStyle]::"Regular")
    for ($i=0; $i -lt $DataGridTabControl.TabCount; $i++) {  # Loop to all Tabs
      if ($i -eq $_.Index) {  # Current Tab: set the colors to highlight the Tab
        $BrushFont=New-Object Drawing.SolidBrush "Black"
        $BrushBack=New-Object Drawing.SolidBrush $DataGridTabControl.TabPages[$i].Tag.ColorSelected
      }
      else {  # Set the colors to non current tabs
        $BrushFont=New-Object Drawing.SolidBrush "Gray"
        $BrushBack=New-Object Drawing.SolidBrush $DataGridTabControl.TabPages[$i].Tag.ColorUnSelected
      }
      $TabX=$DataGridTabControl.GetTabRect($i).X
      $TabY=$DataGridTabControl.GetTabRect($i).Y
      $TabW=$DataGridTabControl.GetTabRect($i).Width
      $DataGridTabControl.CreateGraphics().FillRectangle($BrushBack, $TabX, 1, $TabW-2, 3)  # Paint the line on the top of the Tab
      $DataGridTabControl.CreateGraphics().DrawString($DataGridTabControl.TabPages[$i].Text, $Font, $BrushFont, $TabX+1, 4)  # Write the Tab Label
    }

  }

  $DataGridTabControl_DrawItemHandlerColorsFull={  # Set the Tab Colors and Text
   
    $CurrentTabIndex=$DataGridTabControl.SelectedTab.TabIndex

    if ($($DataGridTabControl.SelectedIndex) -ne $($_.index)) { return }  # Skip useless repaints

    $Font=New-Object Drawing.Font("Microsoft Sans Serif", 8, [Drawing.FontStyle]::"Regular")
    for ($i=0; $i -lt $DataGridTabControl.TabCount; $i++) {  # Loop to all Tabs
      if ($i -eq $_.Index) {  # Current Tab: set the colors to highlight the Tab
        $BrushFont=New-Object Drawing.SolidBrush "Black"
        $BrushBack=New-Object Drawing.SolidBrush $DataGridTabControl.TabPages[$i].Tag.ColorSelected
      }
      else {  # Set the colors to non current tabs
        $BrushFont=New-Object Drawing.SolidBrush "Gray"
        $BrushBack=New-Object Drawing.SolidBrush $DataGridTabControl.TabPages[$i].Tag.ColorUnSelected
      }
      $TabX=$DataGridTabControl.GetTabRect($i).X
      $TabY=$DataGridTabControl.GetTabRect($i).Y
      $TabW=$DataGridTabControl.GetTabRect($i).Width
      $DataGridTabControl.CreateGraphics().FillRectangle($BrushBack, $TabX+1, 0, $TabW-2, 20)  # Paint the Tab
      $DataGridTabControl.CreateGraphics().DrawString($DataGridTabControl.TabPages[$i].Text, $Font, $BrushFont, $TabX+1, 4)  # Write the Tab Label
    }

  }

  $DataGridTabControl.Add_Selected( {  # Set the different handlers for the selected tab/grid
    $Script:OutputDataGrid=$OutputDataGridTab[$DataGridTabControl.SelectedTab.TabIndex]  # Set the main variable $OutputDataGrid to the current grid
    $Script:GridIndex=$DataGridTabControl.SelectedTab.Tag.TabPageIndex
    Get-CountCheckboxes
    $Script:SelectionChanged=$True
    $OutputDataGrid.Add_CellMouseDown( $OutputDataGrid_CellMouseClickHandler )
    $OutputDataGrid.Add_CurrentCellDirtyStateChanged( $OutputDataGrid_CurrentCellDirtyStateChangedHandler )
    $OutputDataGrid.Add_Click( $OutputDataGrid_ClickHandlder )
    $OutputDataGrid.Add_UserAddedRow( $OutputDataGrid_UserAddedRowHandler )
    $OutputDataGrid.Add_UserDeletingRow( $OutputDataGrid_UserDeletingRowHandler )
  })


  $DataGridTabControl_KeyDownHandler=[System.Windows.Forms.KeyEventHandler]{  # Catch the CTRL+V
    try {
      if ($OutputDataGrid.CurrentCell.IsInEditMode) { return }  # A cell is currently beeing edit, exit and use the standard CTRL+V behaviour
    }
    catch {
      return
    }
    $Key=$_
    if ($Key.KeyData -eq "Control, V") { Get-ObjectsPatse }  # Override the CTRL+V action
    if ($Key.KeyData -eq "Delete") {  # Delete the rows when delete is pressed
      try {
        $SelectedColumns=$OutputDataGrid.SelectedCells.ColumnIndex | select -Unique
        if ($SelectedColumns -eq 0) { Set-RightClick_SetNewSelectionFromGrid $False }
      }
      catch { }
    }  
  }
  $DataGridTabControl.Add_KeyDown( $DataGridTabControl_KeyDownHandler )

  $DataGridTabControl_ClickHandler=[System.Windows.Forms.MouseEventHandler]{  # Right click actions on a Tab
    if ($_.Button -eq "Right") {  # Right Click detected   
      for ($i=0; $i -lt $DataGridTabControl.TabCount; $i++) {  
        if ($DataGridTabControl.GetTabRect($i).Contains($_.Location)) {  # Detect the tab corresponding to the click location
          $DataGridTabControlContextMenuObject=New-Object 'System.Windows.Forms.ContextMenuStrip'  # Create the right click context menu
          $DataGridTabControl.ContextMenuStrip=$DataGridTabControlContextMenuObject
          $SBRename=[scriptblock]::Create("Rename-Tab $i")
          Add-ContextMenuStripItem $DataGridTabControlContextMenuObject "Rename the tab $($DataGridTabControl.TabPages[$i].Text)" $SBRename
          if ($DataGridTabControl.TabCount -gt 1) {  # Enable the deleting if more than 1 tab
            Add-ContextStripSeparator $DataGridTabControlContextMenuObject
            Add-ContextMenuStripItem $DataGridTabControlContextMenuObject "Remove the tab" { Remove-Tab }
          }
          if ($TabLook -gt 0 ) {
            Add-ContextStripSeparator $DataGridTabControlContextMenuObject  # Create the Color option
            Add-ContextMenuStripItem $DataGridTabControlContextMenuObject "Tab color" {  }
            $DataGridTabControlContextMenuObjectColor=$DataGridTabControlContextMenuObject.Items | Where { $_.Text -eq "Tab color" } 
            $DataGridTabControlContextMenuObjectColor.DropDownItems.Clear()
            foreach ($color in $TabColorPalette.GetEnumerator()) {  # Loop to the defined colors, and create the submenus
              $SBColor=[scriptblock]::Create("Set-TabColor $($color.Name) $($color.Value)")
              Add-ContextSubMenuStripItem $DataGridTabControlContextMenuObjectColor " " $SBColor $True $color.Name
            }
            return
          }
        }
      }
    }
    else {  # Not a Right Click
      $RunningTask=@($OutputDataGrid.Rows | where { [int]($_.Cells[4].Value) -gt 0 }).Count  # Check if some objects are in a runing state and exits if any
      ($MenuToolStrip.Items | where { $_.ToolTipText -eq "Clear the Grid" }).Enabled=($RunningTask -eq 0)
      $DataGridTabControl.ContextMenuStrip=$Null

      $Script:MouseOnTabBegin=-1  # Save the position of the clicked Tab, for potential Tab permutation
      for ($i=0;$i -le $DataGridTabControl.TabCount - 1; $i++) {  # Check the cursor position
        if (($_.X -gt $DataGridTabControl.GetTabRect($i).Left) -and ($_.X -lt $DataGridTabControl.GetTabRect($i).Right) -and ($_.Y -gt $DataGridTabControl.GetTabRect($i).Top) -and ($_.Y -lt $DataGridTabControl.GetTabRect($i).Bottom)) {  
          $Script:MouseOnTabBegin=$i  # Save the tab position
          $Script:TabMoving=$True  # Enable the TabMoving mode
        }
      }
    }
  }
  $DataGridTabControl.Add_MouseDown( $DataGridTabControl_ClickHandler )

  $Script:TabMoving=$False

  $DataGridTabControl_MouseupHandler=[System.Windows.Forms.MouseEventHandler]{
    $DataGridTabControl.ContextMenuStrip=$Null
    if ($_.Button -eq "Right") {  # Right Click detected 
      return
    }

    $Script:TabMoving=$False  # The mouse is not pressed anymore: disable the TabMoving mode
    $DataGridTabControl.Cursor=[System.Windows.Forms.Cursors]::Default  # Set the cursor to normal

    $MouseOnTab=-1
    for ($i=0;$i -le $DataGridTabControl.TabCount - 1; $i++) {  # Locate the current position of the move and determine the index to use for the permutation
      if (($_.X -gt $DataGridTabControl.GetTabRect($i).Left) -and ($_.X -lt $DataGridTabControl.GetTabRect($i).Right) -and ($_.Y -gt $DataGridTabControl.GetTabRect($i).Top) -and ($_.Y -lt $DataGridTabControl.GetTabRect($i).Bottom)) { 
        if ($_.X -gt $($DataGridTabControl.GetTabRect($i).Left+$DataGridTabControl.GetTabRect($i).Width/2)) {
          if ($i -lt $MouseOnTabBegin) { $MouseOnTab=$i+1 } else { $MouseOnTab=$i }
        }
        else {
          $MouseOnTab=$i
        }
      }
    }

    if (($MouseOnTab -eq -1) -or ($MouseOnTabBegin -eq -1)) { return }  # Nothing to do

    if ($MouseOnTab -gt $MouseOnTabBegin) {  # Move the tabs to right
      for ($i=$MouseOnTabBegin; $i -lt $MouseOnTab; $i++) {
        $Tab1=$DataGridTabControl.TabPages[$i]
        $Tab2=$DataGridTabControl.TabPages[$i+1]
        $DataGridTabControl.TabPages[$i]=$Tab2
        $DataGridTabControl.TabPages[$i+1]=$Tab1
      }
    }

    if ($MouseOnTab -lt $MouseOnTabBegin) {
      for ($i=$MouseOnTabBegin; $i -gt $MouseOnTab; $i--) {  # Move the tabs to left
        $Tab1=$DataGridTabControl.TabPages[$i]
        $Tab2=$DataGridTabControl.TabPages[$i-1]
        $DataGridTabControl.TabPages[$i]=$Tab2
        $DataGridTabControl.TabPages[$i-1]=$Tab1
      }
    }

    $DataGridTabControl.SelectedIndex=$MouseOnTab
    $DataGridTabControl.TabPages[$MouseOnTab].Refresh()
    
  }
  $DataGridTabControl.Add_Mouseup( $DataGridTabControl_MouseupHandler )

  
  $DataGridTabControl_MouseMoveHandler=[System.Windows.Forms.MouseEventHandler]{

    if ($TabMoving -eq $False) { return }

    $Default=$True
    for ($i=0;$i -le $DataGridTabControl.TabCount - 1; $i++) {  # Modify the cursor icon, depending on the position of the mouse
      if (($_.X -gt $DataGridTabControl.GetTabRect($i).Left) -and ($_.X -lt $DataGridTabControl.GetTabRect($i).Right) -and ($_.Y -gt $DataGridTabControl.GetTabRect($i).Top) -and ($_.Y -lt $DataGridTabControl.GetTabRect($i).Bottom)) { 
        $Default=$False
        if ($i -eq $MouseOnTabBegin) { $DataGridTabControl.Cursor=[System.Windows.Forms.Cursors]::Default ; Return }
        if ($_.X -gt $($DataGridTabControl.GetTabRect($i).Left+$DataGridTabControl.GetTabRect($i).Width/2)) {
          $DataGridTabControl.Cursor=[System.Windows.Forms.Cursors]::PanEast
        }
        else {
          $DataGridTabControl.Cursor=[System.Windows.Forms.Cursors]::PanWest
        }
      }
    }

    if ($Default -eq $True) { $DataGridTabControl.Cursor=[System.Windows.Forms.Cursors]::Default }

  }
  $DataGridTabControl.Add_MouseMove( $DataGridTabControl_MouseMoveHandler )


  $DataGridTabControlContextMenuObject=New-Object 'System.Windows.Forms.ContextMenuStrip'
  $DataGridTabControl.ContextMenuStrip=$DataGridTabControlContextMenuObject

  $OutputDataGrid_CellMouseClickHandler=[System.Windows.Forms.DataGridViewCellMouseEventHandler]{  # Right click actions on the grid
    if ($_.Button -ne "Right") { return }  # Not a right click: exit
    try {  # If more than 1 column are selected, or checkboxes: exit
      if ((@($OutputDataGrid.SelectedCells.ColumnIndex | Select -Unique).Count -ne 1) -or (@($OutputDataGrid.SelectedCells.ColumnIndex | Select -Unique) -eq 7)) { 
        $OutputDataGrid.Columns[0].ContextMenuStrip=$Null
        $OutputDataGrid.Columns[2].ContextMenuStrip=$Null
        $OutputDataGrid.Columns[3].ContextMenuStrip=$Null
        $OutputDataGrid.Columns[9].ContextMenuStrip=$Null
        $DataGridTabControl.ContextMenuStrip=$Null
        return 
      }
    }
    catch {  # Nothing selected
      return
    }

    try {  # If the last cells is selected: exit
      if ($OutputDataGrid.SelectedCells.RowIndex -eq $OutputDataGrid.RowCount-1) {
        $OutputDataGrid.Columns[0].ContextMenuStrip=$Null
        $OutputDataGrid.Columns[2].ContextMenuStrip=$Null
        $OutputDataGrid.Columns[3].ContextMenuStrip=$Null
        $OutputDataGrid.Columns[9].ContextMenuStrip=$Null
        $DataGridTabControl.ContextMenuStrip=$Null
      return
      }
    }
    catch {
      $OutputDataGrid.Columns[0].ContextMenuStrip=$Null
      $OutputDataGrid.Columns[2].ContextMenuStrip=$Null
      $OutputDataGrid.Columns[3].ContextMenuStrip=$Null
      $OutputDataGrid.Columns[9].ContextMenuStrip=$Null
      return
    }

    if ($OutputDataGrid.SelectedCells.ColumnIndex -eq 0) {  # The selected cells are in column 0, "Objects"
      $OutputDataGrid.Columns[0].ContextMenuStrip=$OutputDataGridContextMenuObject
      $Script:LoadedFiles=""
      $Script:LoadedFiles=foreach ($RowIndex in $OutputDataGrid.SelectedCells.Rowindex) {  # Search for all files used to load the objects
        if (($OutputDataGrid.rows[$rowindex].Cells[6].Value -ne "") -and ($OutputDataGrid.rows[$rowindex].Cells[6].Value -ne $Null)) { $OutputDataGrid.rows[$rowindex].Cells[6].Value }
      } 
      $Script:LoadedFiles=$LoadedFiles | Select -Unique 
      $OutputDataGridContextMenuObject.Refresh()

      $OutputDataGridContextMenuObject.Items[0].Enabled=($SequenceRunning -eq $False)  # Disable "Set the highlighted objects as new collection" if a sequence is running
      $SeqRun=$False
      foreach ($RowIndex in $OutputDataGrid.SelectedCells.RowIndex) { if ($OutputDataGrid.Rows[$RowIndex].Cells[4].Value -gt 0) { $SeqRun=$True ; break } }
      if (($LoadedFiles -eq $Null) -or ($LoadedFiles -eq "")) {  # No file used to insert the objects
        $OutputDataGridContextMenuObjectRemove.DropDownItems[0].Enabled=!($SeqRun)
        $OutputDataGridContextMenuObjectRemove.DropDownItems[1].Text=""
        $OutputDataGridContextMenuObjectRemove.DropDownItems[1].Visible=$False
        $OutputDataGridContextMenuObjectRemove.DropDownItems[2].Text=""
        $OutputDataGridContextMenuObjectRemove.DropDownItems[2].Visible=$False
      }
      else {  # Some files were used to load objects
        $ListOfFiles=(Split-Path $LoadedFiles -Leaf) -join ", "
        if ($ListOfFiles.Length -gt 30) { $ListOfFiles="$($ListOfFiles.Substring(0,30))..."}
        $OutputDataGridContextMenuObjectRemove.DropDownItems[0].Enabled=!($SeqRun)
        $OutputDataGridContextMenuObjectRemove.DropDownItems[1].Visible=$True
        $OutputDataGridContextMenuObjectRemove.DropDownItems[1].Text="Remove from $ListOfFiles"
        $OutputDataGridContextMenuObjectRemove.DropDownItems[1].Enabled=!($SeqRun)
        $OutputDataGridContextMenuObjectRemove.DropDownItems[2].Visible=$True
        $OutputDataGridContextMenuObjectRemove.DropDownItems[2].Text="Remove from the grid && $ListOfFiles"
        $OutputDataGridContextMenuObjectRemove.DropDownItems[2].Enabled=!($SeqRun)
      }
      
      $OutputDataGridContextMenuObjectTabCopyObject=$OutputDataGridContextMenuObjectTab.DropDownItems | Where { $_.Text -eq "Copy to ..." }
      $OutputDataGridContextMenuObjectTabMoveObject=$OutputDataGridContextMenuObjectTab.DropDownItems | Where { $_.Text -eq "Move to ..." }
      $OutputDataGridContextMenuObjectTabCopyObject.DropDownItems.Clear()
      $OutputDataGridContextMenuObjectTabMoveObject.DropDownItems.Clear()
      Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectTabCopyObject "A new Tab" { Set-CopyObjectsToTab $False -1 }  # Create the "Copy to ... A new Tab"
      Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectTabMoveObject "A new Tab" { Set-CopyObjectsToTab $True -1 }  # Move the "Copy to ... A new Tab"
      if ($DataGridTabControl.TabCount -gt 1) {  # More than one Tab detected 
        Add-ContextSubStripSeparator $OutputDataGridContextMenuObjectTabCopyObject
        Add-ContextSubStripSeparator $OutputDataGridContextMenuObjectTabMoveObject
        foreach ($TabName in $DataGridTabControl.TabPages) {  # Loop through all tabs
          if ($TabName.TabIndex -eq $DataGridTabControl.SelectedTab.TabIndex) { continue }  # Current Tab: exit
          $SBCopy=[scriptblock]::Create("Set-CopyObjectsToTab $False $($TabName.TabIndex)")
          $SBMove=[scriptblock]::Create("Set-CopyObjectsToTab $True $($TabName.TabIndex)")
          Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectTabCopyObject $TabName.Text $SBCopy  # Create a Copy entry for the Tab $Tabname
          Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectTabMoveObject $TabName.Text $SBMove  # Create a Move entry for the Tab $Tabname
        }
      }

      $OutputDataGridContextMenuObjectGroup.DropDownItems.Clear()  # Recreate the Submenu to avoid the multiple Add_Click bug
      Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectGroup "Add to Group" { }
      $OutputDataGridContextMenuObjectAddToGroup=$OutputDataGridContextMenuObjectGroup.DropDownItems | Where { $_.Text -eq "Add to Group" }  
      Add-ContextSubStripSeparator $OutputDataGridContextMenuObjectGroup
      Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectGroup "Assign Sequence" { }
      Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectGroup "Assign Sequence with Scheduler" { }
      Add-ContextSubStripSeparator $OutputDataGridContextMenuObjectGroup
      Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectGroup "Remove objects from Group" { Set-UnAssignSequenceToObjects }

      $GroupValue=$OutputDataGrid.SelectedCells | Select -ExpandProperty Tag | Select -ExpandProperty GroupID -Unique  # Enumerate the groups of the selected objects
      $OutputDataGridContextMenuObjectGroup.DropDownItems[0].Visible=$False
      $OutputDataGridContextMenuObjectGroup.DropDownItems[1].Visible=$False
      $OutputDataGridContextMenuObjectGroup.DropDownItems[4].Visible=$False
      $OutputDataGridContextMenuObjectGroup.DropDownItems[5].Visible=$False
      if (@($GroupValue).Count -ne 1) {  # More than one group found: Disable the "Groups" submenu
        $OutputDataGridContextMenuObject.Items[6].Enabled=$False
        return 
      }
      if ((@($GroupValue).Count -eq 1) -and ($GroupValue -ne "0")) {  # One group found
        $OutputDataGridContextMenuObject.Items[6].Enabled=$True  # Enable the "Groups" submenu
        if ($SequenceLoaded) {  # A sequence is loaded
          $OutputDataGridContextMenuObjectGroup.DropDownItems[2].Visible=$True
          $OutputDataGridContextMenuObjectGroup.DropDownItems[2].Text="Assign '$SequenceName' to Group $GroupValue"
          $OutputDataGridContextMenuObjectGroup.DropDownItems[2].Enabled=$True
          $SBAssign=[scriptblock]::Create("Set-AssignSequenceToObjects $False $GroupValue")
          $OutputDataGridContextMenuObjectGroup.DropDownItems[2].Add_Click( $SBAssign )
          $OutputDataGridContextMenuObjectGroup.DropDownItems[3].Visible=$True
          $OutputDataGridContextMenuObjectGroup.DropDownItems[3].Text="Assign '$SequenceName' with Scheduler to Group $GroupValue"
          $OutputDataGridContextMenuObjectGroup.DropDownItems[3].Enabled=$True
          $SBAssignSched=[scriptblock]::Create("Set-AssignSequenceToObjects $True $GroupValue")
          $OutputDataGridContextMenuObjectGroup.DropDownItems[3].Add_Click( $SBAssignSched )
          $OutputDataGridContextMenuObjectGroup.DropDownItems[4].Visible=$True  # Show the separator
        }
        else {  # No sequence loaded
          $OutputDataGridContextMenuObjectGroup.DropDownItems[2].Visible=$False  # Disable "Assign Sequence to Group"
          $OutputDataGridContextMenuObjectGroup.DropDownItems[3].Visible=$False  # Disable "Assign Sequence with Scheduler to Group"
          $OutputDataGridContextMenuObjectGroup.DropDownItems[4].Visible=$False  # Hide the separator
        }
        $OutputDataGridContextMenuObjectGroup.DropDownItems[5].Visible=$True
        $OutputDataGridContextMenuObjectGroup.DropDownItems[5].Text="Remove from Group $GroupValue"
        $OutputDataGridContextMenuObjectGroup.DropDownItems[5].Enabled=($GroupValue -notin $GroupsRunning)
        return
      }

      if ($SequenceLoaded) {  # No group found in the selected objects ($GroupValue -eq "0")
        $OutputDataGridContextMenuObject.Items[6].Enabled=$True  # Enable the "Groups" submenu and its Assign submenus
        $OutputDataGridContextMenuObjectGroup.DropDownItems[2].Text="Assign the Sequence to a New Group"
        $OutputDataGridContextMenuObjectGroup.DropDownItems[2].Enabled=$True
        $SBAssign=[scriptblock]::Create("Set-AssignSequenceToObjects $False $False")
        $OutputDataGridContextMenuObjectGroup.DropDownItems[2].Add_Click( $SBAssign )
        $OutputDataGridContextMenuObjectGroup.DropDownItems[3].Text="Assign the Sequence with Scheduler to New Group"
        $OutputDataGridContextMenuObjectGroup.DropDownItems[3].Enabled=$True
        $SBAssignSched=[scriptblock]::Create("Set-AssignSequenceToObjects $True $False")
        $OutputDataGridContextMenuObjectGroup.DropDownItems[3].Add_Click( $SBAssignSched )
        $GroupsInCurrentGrid=@(($OutputDataGrid.Rows.Cells | Where { ($_.ColumnIndex -eq 0) } | Select -ExpandProperty Tag) | Where { $_.GroupID -ne 0 } | Select -ExpandProperty GroupID -Unique | Sort-Object)  # Enumerate the groups in the grid
        if ((@($GroupsUsed).Count -gt 0) -and (@($GroupsUsed) -ne "0")) {
          $OutputDataGridContextMenuObjectAddToGroup.DropDownItems.Clear()
          foreach ($GroupUsedItem in $GroupsUsed) {  # Create an entry for each group found
            if (($GroupUsedItem -eq "0") -or ($GroupUsedItem -notin $GroupsInCurrentGrid)) { continue }
            if ($GroupUsedItem -in $GroupsRunning) { continue }
            $SBObjToGrp=[scriptblock]::Create("Set-ObjectsToGroup $GroupUsedItem")
            Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectAddToGroup $GroupUsedItem $SBObjToGrp
          }
          $OutputDataGridContextMenuObjectGroup.DropDownItems[0].Visible=($OutputDataGridContextMenuObjectAddToGroup.DropDownItems.Count -gt 0)  # Enable the "Add to Group" and its separator
          $OutputDataGridContextMenuObjectGroup.DropDownItems[1].Visible=($OutputDataGridContextMenuObjectAddToGroup.DropDownItems.Count -gt 0)
        }
      }
      elseif (@($GroupsUsed).Count -gt 0) {
        $GroupsInCurrentGrid=@(($OutputDataGrid.Rows.Cells | Where { ($_.ColumnIndex -eq 0) } | Select -ExpandProperty Tag) | Where { $_.GroupID -ne 0 } | Select -ExpandProperty GroupID -Unique | Sort-Object)  # Enumerate the groups in the grid
        $OutputDataGridContextMenuObjectAddToGroup.DropDownItems.Clear()
        foreach ($GroupUsedItem in $GroupsUsed) {  # Create an entry for each group found
          if (($GroupUsedItem -eq "0") -or ($GroupUsedItem -notin $GroupsInCurrentGrid)) { continue }
          if ($GroupUsedItem -in $GroupsRunning) { continue }
          $SBObjToGrp=[scriptblock]::Create("Set-ObjectsToGroup $GroupUsedItem")
          Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectAddToGroup $GroupUsedItem $SBObjToGrp
        }
        $OutputDataGridContextMenuObjectGroup.DropDownItems[0].Visible=$True
        $OutputDataGridContextMenuObjectGroup.DropDownItems[2].Text=""
        $OutputDataGridContextMenuObjectGroup.DropDownItems[2].Visible=$False
        $OutputDataGridContextMenuObjectGroup.DropDownItems[3].Text=""
        $OutputDataGridContextMenuObjectGroup.DropDownItems[3].Visible=$False
      }
      else {
        $OutputDataGridContextMenuObject.Items[6].Enabled=$False  # Disable the "Group" menu
      }

    }

    if ($OutputDataGrid.SelectedCells.ColumnIndex -eq 2) {  # The selected cells are in column 2, "Tasks Results"
      $OutputDataGrid.Columns[2].ContextMenuStrip=$OutputDataGridContextMenuResult
      $OutputDataGridContextMenuResult.Items[0].Visible=$False  # Hide "Show Variable" per default
      $OutputDataGridContextMenuResult.Items[1].Visible=$False  # Hide the corresponding separator per default
      $SeqIdFound=foreach ($item in $OutputDataGrid.SelectedCells) { $OutputDataGrid.rows[$item.RowIndex].Cells[8].Value }  # Enumerate the sequence ID's of the cells selected
      $SeqIdFound=$($SeqIdFound | Select -Unique)
      if ((@($SeqIdFound).Count -eq 1) -and ($SeqIdFound -gt 0)) {  # Only one sequence ID found
        $VarText=@() 
        for ($i=1; $i -le $nbVariableTypes; $i++) {  # Loop into the variable types
          $Sequences[$SeqIdFound].ScriptBlockVariable[$i].GetEnumerator() | ForEach-Object { $VarText+="$($_.Key) : $($_.Value)" }  # Extend $VarText with the variables names and values
        } 
        if ($VarText -ne "") {  # Some variables were found
          $OutputDataGridContextMenuResult.Items[0].Visible=$True  # Show "Show Variable" per default
          $OutputDataGridContextMenuResult.Items[1].Visible=$True  # Hide the corresponding separator
          $OutputDataGridContextMenuResultVariable.DropDownItems[0].Text=$VarText -join "`r`n"  # Show the variables as submenu
        }
      }
      Set-RightClick_ShowProtocol  # Creates the steps protocol
    }

    if ($OutputDataGrid.SelectedCells.ColumnIndex -eq 3) {  # The selected cells are in column 3, "State"
      if ($SequenceRunning -eq $True) {  # Some sequences are running
        $OutputDataGrid.Columns[3].ContextMenuStrip=$OutputDataGridContextMenuStateRunning  # Display the "Cancel" context
      }
      else {
        $OutputDataGrid.Columns[3].ContextMenuStrip=$OutputDataGridContextMenuStatePending  # Display the "State" context
        $JobRes=$OutputDataGrid.SelectedCells.Value | Select -Unique
        if (($JobRes | Measure-Object).Count -gt 1) {  # More than one State selected
          $OutputDataGridContextMenuStatePending.Items[0].Text="Select only one kind of result states" 
          $OutputDataGridContextMenuStatePending.Items[1].Visible=$False 
        }
        else {
          $OutputDataGridContextMenuStatePending.Items[0].Text="Highlight all objects in the state $JobRes"
          $OutputDataGridContextMenuStatePending.Items[1].Text="Remove all objects in the state $JobRes"
          $OutputDataGridContextMenuStatePending.Items[1].Visible=$True
          $Script:SelectedState=$JobRes
        }
      }
    }

    if ($OutputDataGrid.SelectedCells.ColumnIndex -eq 9) {  # The selected cells are in column 9, "Group"
      $Script:GroupsFoundForMaxThreads=foreach ($RowIndex in $OutputDataGrid.SelectedCells.Rowindex) {  # Select the Groups selected
        $OutputDataGrid.rows[$rowindex].Cells[0].Tag.GroupID
      } 
      $Script:GroupsFoundForMaxThreads=$GroupsFoundForMaxThreads | Select -Unique
      if ((@($GroupsFoundForMaxThreads).Count -ne 1) -or ($GroupsFoundForMaxThreads -le 0)) {  # Not only one group selected
        $OutputDataGridContextMenuGroup.Items[0].Text="Wrong selection"
        $OutputDataGridContextMenuGroup.Items[0].Enabled=$False
      }
      else {
        $OutputDataGridContextMenuGroup.Items[0].Text="Change Max. Threads of Group $GroupsFoundForMaxThreads"
        $OutputDataGridContextMenuGroup.Items[0].Enabled=($GroupsFoundForMaxThreads -notin $GroupsRunning)  # Enable if the Group selected is not currently running
      }
      $OutputDataGrid.Columns[9].ContextMenuStrip=$OutputDataGridContextMenuGroup
    }
  }
  $OutputDataGrid.Add_CellMouseDown( $OutputDataGrid_CellMouseClickHandler )

  $OutputDataGrid_CurrentCellDirtyStateChangedHandler= { 
    if ($OutputDataGrid.CurrentCell.ColumnIndex -eq 7) { 
      $OutputDataGrid.EndEdit()
      if ($OutputDataGrid.CurrentCell.Value -eq $True) { 
        $Script:nbCheckedBoxes=$Script:nbCheckedBoxes+0.5 
      } else { 
        $Script:nbCheckedBoxes=$Script:nbCheckedBoxes-0.5 
      } 
      Get-CountCheckboxes
    }
  }
  $OutputDataGrid.Add_CurrentCellDirtyStateChanged( $OutputDataGrid_CurrentCellDirtyStateChangedHandler )

  $OutputDataGrid_ClickHandlder={ 
    try {  # Set the last checkbox to Read Only when a user tries to click it
      if (($OutputDataGrid.SelectedCells.RowIndex -eq $($OutputDataGrid.RowCount -1)) -and ($OutputDataGrid.SelectedCells.ColumnIndex -eq 7)) { 
        $OutputDataGrid.Rows[$($OutputDataGrid.RowCount -1)].Cells[7].ReadOnly=$True
      }
    }
    catch {  # The user has clicked outside the grid
      return
    }
  } 
  $OutputDataGrid.Add_Click( $OutputDataGrid_ClickHandlder )

  $OutputDataGrid_CellEndEditHandler= {
    try {  # Remove the row if the object name has been set to blank
      $ObjectName=$OutputDataGrid.Rows[$OutputDataGrid.SelectedCells.RowIndex].Cells[0].Value
      if ([string]::IsNullOrWhiteSpace($ObjectName)) {  # The object cell is empty
        $OutputDataGrid.Rows.RemoveAt($OutputDataGrid.SelectedCells.RowIndex)  # Remove the row
        Get-CountCheckboxes
      }
    }
    catch {
      return
    }
  }
  $OutputDataGrid.Add_CellEndEdit( $OutputDataGrid_CellEndEditHandler )

  $OutputDataGrid_UserAddedRowHandler= {  # New row added
    Set-CellValue $GridIndex $OutputDataGrid.CurrentCell.RowIndex 0 "Pending" "Pending" 0 "#" "#" 0  # Default values
    $OutputDataGrid.Rows[$OutputDataGrid.CurrentCell.RowIndex].Cells[7].ReadOnly=$False  # Enable the checkbox
    $OutputDataGrid.Rows[$OutputDataGrid.CurrentCell.RowIndex].Cells[7].Value=$True  # Check the checkbox
    $OutputDataGrid.Rows[$OutputDataGrid.CurrentCell.RowIndex].Cells[9].Value="-"  # Comment for the column "Group"
    $ObjectOptions=New-Object -TypeName PSObject
    $ObjectOptions | Add-Member –MemberType NoteProperty –Name GroupID –Value "0"
    $ObjectOptions | Add-Member –MemberType NoteProperty –Name PreviousStateComment –Value ""
    $ObjectOptions | Add-Member –MemberType NoteProperty –Name StepProtocol –Value $Null
    $ObjectOptions | Add-Member –MemberType NoteProperty –Name SharedVariable –Value $Null
    $OutputDataGrid.Rows[$OutputDataGrid.CurrentCell.RowIndex].Cells[0].Tag=$ObjectOptions
    Get-CountCheckboxes 
  }
  $OutputDataGrid.Add_UserAddedRow( $OutputDataGrid_UserAddedRowHandler )
  
  $OutputDataGrid_UserDeletingRowHandler=[System.Windows.Forms.DataGridViewRowCancelEventHandler ]{  # Rows have to be deleted: cancel the operation and delete with Set-RightClick_SetNewSelectionFromGrid
    if ($OutputDataGrid.SelectedCells.Count -eq 0) { $_.Cancel=$True ; return }  # No object selected: cancel the deletion
    $IsRunningSeqSelected=@($OutputDataGrid.SelectedRows | where { $_.Cells[4].Value -gt 0 }).Count
    if ($IsRunningSeqSelected -gt 0)  {  # Trying to delete objects running a sequence: cancel the deletion
      $_.Cancel=$True
      return 
    }
    try {
      $RowsSelectedIndex=$OutputDataGrid.SelectedRows.Index  # Enumerate the rows to deleted
    }
    catch {
      $RowsSelectedIndex=$OutputDataGrid.SelectedCells.RowIndex
    }
    $_.Cancel=$True  # Cancel the deletion
    $OutputDataGrid.ClearSelection()
    foreach ($Index in $RowsSelectedIndex) {  # Loop into the rows selected
      if (($OutputDataGrid.Rows[$Index].Cells[4].Value -le 0) -and ($OutputDataGrid.Rows[$Index].Cells[4].Value -ne 5)) {  # Select the objects not running
        $OutputDataGrid.Rows[$Index].Cells[0].Selected=$True 
      }   
    }
    Set-RightClick_SetNewSelectionFromGrid $False  # Delete the selected objects
  }
  $OutputDataGrid.Add_UserDeletingRow( $OutputDataGrid_UserDeletingRowHandler )

  # Define the context menu for the Object column

  $OutputDataGridContextMenuObject=New-Object 'System.Windows.Forms.ContextMenuStrip'
  $OutputDataGrid.Columns[0].ContextMenuStrip=$OutputDataGridContextMenuObject
  Add-ContextMenuStripItem $OutputDataGridContextMenuObject "Set the highlighted objects as new collection" { 
    $ReloadVariables=[System.Windows.Forms.MessageBox]::Show("This will remove all other objects from the grid.`r`nDo you really want to perform this operation ?", "WARNING", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($ReloadVariables -eq "yes") {
      Set-RightClick_SetNewSelectionFromGrid $True 
    }
  }
  Add-ContextMenuStripItem $OutputDataGridContextMenuObject "Select the highlighted objects" { Set-RightClick_CheckObject $True }
  Add-ContextMenuStripItem $OutputDataGridContextMenuObject "De-select the highlighted objects" { Set-RightClick_CheckObject $False }
  Add-ContextStripSeparator $OutputDataGridContextMenuObject
  Add-ContextMenuStripItem $OutputDataGridContextMenuObject "Remove the highlighted objects ..." { }
  $OutputDataGridContextMenuObjectRemove=$OutputDataGridContextMenuObject.Items | Where { $_.Text -eq "Remove the highlighted objects ..." }
  Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectRemove "Remove from the grid" { Set-RightClick_SetNewSelectionFromGrid $False }
  Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectRemove "Remove from File Only" { Set-RightClick_RemoveSelectionFromFiles $LoadedFiles $False }
  Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectRemove "Remove from File && Grid" { Set-RightClick_RemoveSelectionFromFiles $LoadedFiles $True }
  Add-ContextStripSeparator $OutputDataGridContextMenuObject

  Add-ContextMenuStripItem $OutputDataGridContextMenuObject "Group" { }
  $OutputDataGridContextMenuObjectGroup=$OutputDataGridContextMenuObject.Items | Where { $_.Text -eq "Group" }  
  Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectGroup "Add to Group" { }
  $OutputDataGridContextMenuObjectAddToGroup=$OutputDataGridContextMenuObjectGroup.DropDownItems | Where { $_.Text -eq "Add to Group" }  
  Add-ContextSubStripSeparator $OutputDataGridContextMenuObjectGroup
  Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectGroup "Assign Sequence" { }
  Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectGroup "Assign Sequence with Scheduler" { }
  Add-ContextSubStripSeparator $OutputDataGridContextMenuObjectGroup
  Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectGroup "Remove objects from Group" { Set-UnAssignSequenceToObjects }
  Add-ContextStripSeparator $OutputDataGridContextMenuObject

  Add-ContextMenuStripItem $OutputDataGridContextMenuObject "Copy or Move the Objects to another Tab" { }
  $OutputDataGridContextMenuObjectTab=$OutputDataGridContextMenuObject.Items | Where { $_.Text -eq "Copy or Move the Objects to another Tab" } 
  Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectTab "Copy to ..." { }
  Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectTab "Move to ..." { }

  # Define the context menu for the Group column
  $OutputDataGridContextMenuGroup=New-Object 'System.Windows.Forms.ContextMenuStrip'
  Add-ContextMenuStripItem $OutputDataGridContextMenuGroup "Change Max. Threads" { Set-GroupMaxThreads }

  # Define the context menu for the Task Result column
  $OutputDataGridContextMenuResult=New-Object 'System.Windows.Forms.ContextMenuStrip'
  $OutputDataGrid.Columns[2].ContextMenuStrip=$OutputDataGridContextMenuResult
  Add-ContextMenuStripItem $OutputDataGridContextMenuResult "Show Variables" { }
  $OutputDataGridContextMenuResultVariable=$OutputDataGridContextMenuResult.Items | Where { $_.Text -eq "Show Variables" }  
  Add-ContextSubMenuStripItem $OutputDataGridContextMenuResultVariable "Show Variables" { }
  Add-ContextStripSeparator $OutputDataGridContextMenuResult
  Add-ContextMenuStripItem $OutputDataGridContextMenuResult "Show Protocol" { }
  $OutputDataGridContextMenuObjectProtocol=$OutputDataGridContextMenuResult.Items | Where { $_.Text -eq "Show Protocol" }  
  Add-ContextSubMenuStripItem $OutputDataGridContextMenuObjectProtocol "Show Protocol" { }
  Add-ContextStripSeparator $OutputDataGridContextMenuResult
  Add-ContextMenuStripItem $OutputDataGridContextMenuResult "Show Sequence Steps" { Show-SequenceSteps }

  # Define the context menus for the State column
  $OutputDataGridContextMenuStateRunning=New-Object 'System.Windows.Forms.ContextMenuStrip'
  Add-ContextMenuStripItem $OutputDataGridContextMenuStateRunning "Cancel" { Cancel-Sequence $OutputDataGrid.SelectedCells.RowIndex }

  $OutputDataGridContextMenuStatePending=New-Object 'System.Windows.Forms.ContextMenuStrip'
  $OutputDataGrid.Columns[3].ContextMenuStrip=$OutputDataGridContextMenuStatePending
  Add-ContextMenuStripItem $OutputDataGridContextMenuStatePending "Highlight all objects in the state" { Set-RightClick_SetNewSelectionFromState $SelectedState $False }
  Add-ContextMenuStripItem $OutputDataGridContextMenuStatePending "Remove all objects in the state" { Set-RightClick_SetNewSelectionFromState $SelectedState $True }

  # Bottom Panel 

  $PanelBottom=New-Object System.Windows.Forms.Panel
  $PanelBottom.Anchor=14
  $PanelBottom.Left=10
  $PanelBottom.Top=$PanelBottomTop
  $PanelBottom.Name="PanelBottom"
  $PanelBottom.Width=$PanelBottomWidth
  $PanelBottom.Height=80
  [void]$SplitContainer1.Panel2.Controls.Add($PanelBottom)

  $ActionButton=New-Object System.Windows.Forms.Button
  $ActionButton.Anchor=6
  $ActionButton.Location="10,8"
  $ActionButton.Name="ActionButton"
  $ActionButton.Size="65,65"
  $ActionButton.Image=$Icon00
  $ActionButton.FlatStyle="Flat"
  $ActionButton.BackColor="Transparent"
  $ActionButton.FlatAppearance.MouseDownBackColor=[System.Drawing.Color]::FromArgb(0,255,255,255)
  $ActionButton.FlatAppearance.MouseOverBackColor=[System.Drawing.Color]::FromArgb(0,255,255,255)
  $ActionButton.FlatAppearance.BorderSize=0
  $ActionButton.Enabled=$False
  $ActionButton.Add_Click( { 
    Start-Sequence
  } )
  [void]$PanelBottom.Controls.Add($ActionButton)

  $ThreadsGroupBox=New-Object System.Windows.Forms.GroupBox
  $ThreadsGroupBox.Location="104,8"
  $ThreadsGroupBox.Size="143,62"
  $ThreadsGroupBox.Text="Threads"
  $PanelBottom.Controls.Add($ThreadsGroupBox)

  $ThreadsLabel=New-Object System.Windows.Forms.Label
  $ThreadsLabel.Location="15,25"
  $ThreadsLabel.Size="74,23"
  $ThreadsLabel.Text="Max. Threads"
  $ThreadsGroupBox.Controls.Add($ThreadsLabel)

  $MaxThreadsText=New-Object System.Windows.Forms.TextBox
  $MaxThreadsText.Location="95,22"
  $MaxThreadsText.Size="28,20"
  $MaxThreadsText.Text="5"
  $MaxThreadsText.TextAlign=2
  $MaxThreadsText.TabIndex=4
  $ThreadsGroupBox.Controls.Add($MaxThreadsText)
  $MaxThreadsText.Add_Leave( { $Script:SelectionChanged=$True } )
  $MaxThreadsText.Add_TextChanged({  
    $This.Text=$This.Text -replace '\D'  # Accept only digits
    $This.Select($This.Text.Length, 0)  # Avoid the cursor jumping back to start
  })

  $ObjectsGroupBox=New-Object System.Windows.Forms.GroupBox
  $ObjectsGroupBox.Location="265,8"
  $ObjectsGroupBox.Size="300,61"
  $ObjectsGroupBox.Text="Objects"
  $PanelBottom.Controls.Add($ObjectsGroupBox)

  $ObjectsLabel=New-Object System.Windows.Forms.Label
  $ObjectsLabel.Location ="16,25"
  $ObjectsLabel.Size="280,23"
  $ObjectsLabel.Text=""
  $ObjectsGroupBox.Controls.Add($ObjectsLabel)

  # Icons Menu Strip

  $MenuToolStrip=New-Object System.Windows.Forms.ToolStrip
  $MenuToolStrip.SuspendLayout()
  $MenuToolStrip.AutoSize=$False
  $MenuToolStrip.Height=38
  $MenuToolStrip.Location="0,40"
  $MenuToolStrip.ImageScalingSize="32,32"
  $MenuToolStrip.RenderMode="System"
  $MenuToolStrip.GripStyle="Hidden"
  [void]$Form.Controls.Add($MenuToolStrip)

  Add-ToolStripButton "Load Objects File" $Icon30 { Get-ObjectsFile } $MenuToolStrip 36
  Add-ToolStripButton "Enter Objects Manually" $Icon31 { Get-ObjectsManually } $MenuToolStrip 36
  Add-ToolStripButton "Paste Objects" $Icon32 { Get-ObjectsPatse } $MenuToolStrip 36
  Add-ToolStripButton "Query AD" $Icon33 { $FormADQuery.ShowDialog() } $MenuToolStrip 36
  Add-ToolStripButton "Query SCCM" $Icon34 { $FormSCCMQuery.ShowDialog() } $MenuToolStrip 36
  Add-ToolStripSeparator  
  Add-ToolStripButton "Load Sequence Manually" $Icon35 { Get-SequenceFileManual } $MenuToolStrip 36
  Add-ToolStripSeparator 
  Add-ToolStripButton "Export to CSV" $Icon36 { $FormExportColorCheckBox.Checked=$False ; $FormExportColorCheckBox.Enabled=$False ; $Script:ExportFormat=0 ; $FormExport.ShowDialog() } $MenuToolStrip 36
  Add-ToolStripButton "Export to Excel" $Icon37 { $FormExportColorCheckBox.Checked=$True ; $FormExportColorCheckBox.Enabled=$True ; $Script:ExportFormat=1 ; $FormExport.ShowDialog() } $MenuToolStrip 36
  Add-ToolStripButton "Export to HTML" $Icon38 { $FormExportColorCheckBox.Checked=$True ; $FormExportColorCheckBox.Enabled=$True ; $Script:ExportFormat=2 ; $FormExport.ShowDialog() } $MenuToolStrip 36
  Add-ToolStripButton "Send to Mail" $Icon39 { $FormExportColorCheckBox.Checked=$True ; $FormExportColorCheckBox.Enabled=$True ; $Script:ExportFormat=3 ; $FormExport.ShowDialog() } $MenuToolStrip 36
  Add-ToolStripSeparator
  Add-ToolStripButton "New Tab" $Icon40 { Set-NewTab ; $DataGridTabControl.SelectedIndex=$DataGridTabControl.TabCount-1 ; $OutputDataGrid.ClearSelection() } $MenuToolStrip 36
  Add-ToolStripSeparator
  Add-ToolStripButton "Select All Objects" $Icon41 { Set-CheckAll $True } $MenuToolStrip 36
  Add-ToolStripButton "De-select All Objects" $Icon42 { Set-CheckAll $False } $MenuToolStrip 36
  Add-ToolStripButton "Clear the Grid" $Icon43 { Clear-Grid } $MenuToolStrip 36
  Add-ToolStripSeparator
  Add-ToolStripButton "Cancel All" $Icon44 { Cancel-Sequence @(0..$($OutputDataGrid.RowCount-2)) } $MenuToolStrip 36
  ($MenuToolStrip.Items | where { $_.ToolTipText -eq "Cancel All" }).Enabled=$False

  # Main menu definition

  $MenuMain=New-Object System.Windows.Forms.MenuStrip
  [void]$Form.Controls.Add($MenuMain) 
  Add-MenuItem "Load"  ([ordered]@{'Load from a file'={ Get-ObjectsFile } ; 
                                   'Load manually'={ Get-ObjectsManually } ; 
                                   'Paste Objects'={ Get-ObjectsPatse } ; 
                                   'Query AD'={ $FormADQuery.ShowDialog() } ; 
                                   'Query SCCM'={ $FormSCCMQuery.ShowDialog() };
                                   'Create an IP range'={ $FormIPRange.ShowDialog() } })
  Set-SubMenuIcons "Load" @($Icon100, $Icon101, $Icon102, $Icon103, $Icon104, $Icon134)
  Add-MenuItem "Sequence" ([ordered]@{'Load a Sequence manually'={ Get-SequenceFileManual } ;
                                      'Separator1'='-' ;
                                      'Reload the Sequence List'={ Set-ReloadSequenceList } ;
                                      'Load a new Sequence List'={ Get-NewSequenceList } })
  Set-SubMenuIcons "Sequence" @($Icon105, $Null, $Icon114, $Icon115)
  Add-MenuItem "Export"   ([ordered]@{'Export to CSV'={ $FormExportColorCheckBox.Checked=$False ; $FormExportColorCheckBox.Enabled=$False ; $Script:ExportFormat=0 ; $FormExport.ShowDialog() } ; 
                                      'Export to Excel'={ $FormExportColorCheckBox.Checked=$True ; $FormExportColorCheckBox.Enabled=$True ; $Script:ExportFormat=1 ; $FormExport.ShowDialog() } ; 
                                      'Export to HTML'={ $FormExportColorCheckBox.Checked=$True ; $FormExportColorCheckBox.Enabled=$True ; $Script:ExportFormat=2 ; $FormExport.ShowDialog() } ;
                                      'Separator1'='-' ;
                                      'Send to Mail'={ $FormExportColorCheckBox.Checked=$True ; $FormExportColorCheckBox.Enabled=$True ; $Script:ExportFormat=3 ; $FormExport.ShowDialog() } })
  Set-SubMenuIcons "Export" @($Icon106, $Icon107, $Icon108, $Null, $Icon109)
  Add-MenuItem "Objects"     ([ordered]@{'Select All Objects'={ Set-CheckAll $True } ;
                                         'De-select All Objects'={ Set-CheckAll $False } ;
                                         'Separator1'='-' ;
                                         'Reset All Objects'={ Reset-AllObjects } ;
                                         'Separator2'='-' ;
                                         'Clear Grid'={ Clear-Grid } })
  Set-SubMenuIcons "Objects" @($Icon110, $Icon111, $Null, $Icon116, $Null, $Icon112)
  Add-MenuItem "Groups"    ([ordered]@{'Import Group or Group List'={ Import-Group } ;
                                       'Export Group'={ } ;
                                       'Separator1'='-' ;
                                       'Select All Objects in Group'={ } ;
                                       'De-select All Objects in Group'={ } })
  Set-SubMenuIcons "Groups" @($Icon117, $Icon118, $Null, $Icon119, $Icon120)
  Add-MenuItem "Tabs"     ([ordered]@{'Add a new Tab'={ Set-NewTab ; $DataGridTabControl.SelectedIndex=$DataGridTabControl.TabCount-1 ; $OutputDataGrid.ClearSelection() } ;
                                      'Remove Tab'={ Remove-Tab } ;
                                      'Separator1'='-' ;
                                      'Load Tabs'={ Import-Tabs } ;
                                      'Save the Tabs'={ Export-Tabs } })
  Set-SubMenuIcons "Tabs" @($Icon128, $Icon129, $Null, $Icon132, $Icon133)
  Add-MenuItem "View"     ([ordered]@{'Automatically set the columns size'={ Set-View_ColumnsSizeAuto }  ;
                                      'Manually set the columns size'={ Set-View_ColumnsSizeManual } ;
                                      'Wrap Text'={ Set-View_Wrap } })
  Set-SubMenuIcons "View" @($Icon121, $Icon122)
  ($MenuMain.Items.DropDown.Items | where { $_.Text -eq "Wrap Text" }).Checked=$False
  Add-MenuItem "Settings" ([ordered]@{'Settings'={ $FormSettings.ShowDialog() } ;
                                      'Reset to defaults'={ Reset-DefaultSettings } })
  Set-SubMenuIcons "Settings" @($Icon124, $Icon125)
  Add-MenuItem "Cancel"   ([ordered]@{'Cancel the Selection'={ Cancel-Sequence $OutputDataGrid.SelectedCells.RowIndex } ;
                                      'Cancel All'={ Cancel-Sequence @(0..$($OutputDataGrid.RowCount-2)) } ;
                                      'Separator1'='-' ;
                                      'Cancel All (No Wait)'={ Cancel-AllForce } })
  Set-SubMenuIcons "Cancel" @($Icon130, $Icon113, $Null, $Icon131)
  ($MenuMain.Items | where { $_.Text -eq "Cancel" }).Enabled=$False
  Add-MenuItem "Help"     ([ordered]@{'User Guide'={ Start-Process "$HydraDocsPath\Hydra_UserGuide.pdf" } ;
                                      'Developer Guide'={ Start-Process "$HydraDocsPath\Hydra_DeveloperGuide.pdf" } ;
                                      'About'={ $FormAbout.ShowDialog() } })
  Set-SubMenuIcons "Help" @($Icon126, $Icon127, $HydraIcon)

  # Main menu mouse actions

  ($MenuMain.Items | where { $_.Text -eq "Tabs" }).Add_MouseHover( {  # Disable the "Remove Tab" if there is only one tab
    ($MenuMain.Items.DropDown.Items | where { $_.Text -eq "Remove Tab" }).Enabled=($DataGridTabControl.TabCount -gt 1)
  })

  ($MenuMain.Items | where { $_.Text -eq "Groups" }).Add_MouseHover( {  # Dynamicaly create Groups submenus
    $MenuItem=New-Object System.Windows.Forms.ToolStripMenuItem
    [void]$ExportGroupMenu.DropDownItems.Add($MenuItem)
    $MenuItem=New-Object System.Windows.Forms.ToolStripMenuItem
    [void]$SelectAllGroupMenu.DropDownItems.Add($MenuItem)
    $MenuItem=New-Object System.Windows.Forms.ToolStripMenuItem
    [void]$DeSelectAllGroupMenu.DropDownItems.Add($MenuItem)
  })

  $ExportGroupMenu=($MenuMain.Items.DropDown.Items | where { $_.Text -eq "Export Group" })
  $ExportGroupMenu_MouseHoverHanlder={  # Create dynamic entries in the Export Group submenu
    $ExportGroupMenu.DropDownItems.Clear()  # Clear all entries and recreate them based on all groups of the current grid
    $nbOfGrp=0
    $GroupsInCurrentGrid=@(($OutputDataGrid.Rows.Cells | Where { ($_.ColumnIndex -eq 0) } | Select -ExpandProperty Tag) | Where { $_.GroupID -ne 0 } | Select -ExpandProperty GroupID -Unique | Sort-Object)  # Enumerate the groups in the grid
    foreach ($GroupUsedItem in $GroupsInCurrentGrid) {  # Create one entry per group
      if ($GroupUsedItem -eq "0") { continue }  # Escape the Group 0 (no group)
      $ExportGroupSB=[scriptblock]::Create("Export-Group $GroupUsedItem")
      Add-ContextSubMenuStripItem $ExportGroupMenu $GroupUsedItem $ExportGroupSB
      $nbOfGrp++
    }
    if ($nbOfGrp -ge 2) {  # If more than 2 groups were found, add the "All Groups" entry
      Add-ContextSubStripSeparator $ExportGroupMenu
      $ExportGroupSB=[scriptblock]::Create("Export-Group 'All Groups'")
      Add-ContextSubMenuStripItem $ExportGroupMenu "All Groups" $ExportGroupSB
    }
  }
  $ExportGroupMenu.Add_MouseHover( $ExportGroupMenu_MouseHoverHanlder )

  $SelectAllGroupMenu=($MenuMain.Items.DropDown.Items | where { $_.Text -eq "Select All Objects in Group" })  # "Select All Objects in Group" operations
  $SelectAllGroupMenu_MouseHoverHanlder={
    $SelectAllGroupMenu.DropDownItems.Clear()  # Clear all entries and recreate them based on all groups of the current grid
    $nbOfGrp=0
    foreach ($GroupUsedItem in $GroupsUsed) {   # Create dynamic entries in the Select Group submenu
      if ($GroupUsedItem -eq "0") { continue }  # Escape the Group 0 (no group)
      $SelectGroupSB=[scriptblock]::Create("Set-SelectGroup $GroupUsedItem $True")
      Add-ContextSubMenuStripItem $SelectAllGroupMenu $GroupUsedItem $SelectGroupSB
      $nbOfGrp++
    }
    if ($nbOfGrp -ge 2) {  # If more than 2 groups were found, add the "All Groups" entry
      Add-ContextSubStripSeparator $SelectAllGroupMenu
      $SelectGroupSB=[scriptblock]::Create("Set-SelectGroup 'All Groups' $True")
      Add-ContextSubMenuStripItem $SelectAllGroupMenu $GroupUsedItem $SelectGroupSB
    }
  }
  $SelectAllGroupMenu.Add_MouseHover( $SelectAllGroupMenu_MouseHoverHanlder )

  $DeSelectAllGroupMenu=($MenuMain.Items.DropDown.Items | where { $_.Text -eq "De-select All Objects in Group" })  # "De-Select All Objects in Group" operations
  $DeSelectAllGroupMenu_MouseHoverHanlder={
    $DeSelectAllGroupMenu.DropDownItems.Clear()  # Clear all entries and recreate them based on all groups of the current grid
    $nbOfGrp=0
    foreach ($GroupUsedItem in $GroupsUsed) {  # Create dynamic entries in the De-Select Group submenu
      if ($GroupUsedItem -eq "0") { continue }  # Escape the Group 0 (no group)
      $SelectGroupSB=[scriptblock]::Create("Set-SelectGroup $GroupUsedItem $False")
      Add-ContextSubMenuStripItem $DeSelectAllGroupMenu $GroupUsedItem $SelectGroupSB
      $nbOfGrp++
    }
    if ($nbOfGrp -ge 2) {  # If more than 2 groups were found, add the "All Groups" entry
      Add-ContextSubStripSeparator $DeSelectAllGroupMenu
      $SelectGroupSB=[scriptblock]::Create("Set-SelectGroup 'All Groups' $False")
      Add-ContextSubMenuStripItem $DeSelectAllGroupMenu $GroupUsedItem $SelectGroupSB
    }
  }
  $DeSelectAllGroupMenu.Add_MouseHover( $DeSelectAllGroupMenu_MouseHoverHanlder )

  # SCCM query window

  $FormSCCMQuery=New-Object System.Windows.Forms.Form
  $FormSCCMQuery.Size='350, 480'
  $FormSCCMQuery.FormBorderStyle="FixedDialog"
  $FormSCCMQuery.StartPosition="CenterParent"
  $FormSCCMQuery.BackColor=$ColorBackground
  $FormSCCMQuery.Add_Load( { $FormSCCMQuery.BackColor=$ColorBackground } )
  $FormSCCMQuery.MinimizeBox=$False
  $FormSCCMQuery.MaximizeBox=$False
  $FormSCCMQuery.WindowState="Normal"
  $FormSCCMQuery.SizeGripStyle="Hide"
  $FormSCCMQuery.FormBorderStyle="FixedDialog"
  $FormSCCMQuery.KeyPreview=$True
  $FormSCCMQuery.Add_KeyDown( {
    if ($_.KeyCode -eq "Escape") { $FormSCCMQuery.Close() } 
  } )
  $FormSCCMQuery.Visible=$False
  $FormSCCMQuery.Text="SCCM Query"

  $SCCMGroupBox=New-Object System.Windows.Forms.GroupBox
  $SCCMGroupBox.Location=New-Object System.Drawing.Size(15,15)
  $SCCMGroupBox.Size=New-Object System.Drawing.Size(320,120)
  $SCCMGroupBox.BackColor="Transparent"
  $SCCMGroupBox.text="SCCM Options"
  $FormSCCMQuery.Controls.Add($SCCMGroupBox)

  $SCCMCountryLabel=New-Object System.Windows.Forms.Label
  $SCCMCountryLabel.Text="Country:"
  $SCCMCountryLabel.Location=New-Object System.Drawing.Size(15,25)
  $SCCMCountryLabel.Size=New-Object System.Drawing.Size(80,20)
  $SCCMGroupBox.Controls.Add($SCCMCountryLabel)

  $SCCMCountryComboBox=New-Object System.Windows.Forms.ComboBox
  $SCCMCountryComboBox.Location=New-Object System.Drawing.Point(100, 22)
  $SCCMCountryComboBox.Size=New-Object System.Drawing.Size(100, 310)
  $SCCMCountryComboBox.AutoCompleteMode='SuggestAppend'
  $SCCMCountryComboBox.AutoCompleteSource='ListItems'
  $SCCMCountryComboBox.DropDownStyle='DropDownList'
  if (Test-Path $CountriesList) {  # Load the countries
    $CountryNames=Import-Csv -Delimiter ";" -Path $CountriesList -Header Country, Server, SiteCode
    foreach ($Elem in $CountryNames) {
     [void]$SCCMCountryComboBox.Items.Add($Elem.Country)  # Add the countries in the combobox
    }
    $SCCMCountryComboBox.SelectedIndex=0
  }
  else {  # No countries file found
    $SCCMCountryComboBox.Enabled=$False
  }
  $SCCMCountryComboBox.Add_SelectedIndexChanged( {  # Set the name and the site code of SCCM depending on the country clicked
    $SCCMServerText.Text=($CountryNames | Where { $_.Country -eq $SCCMCountryComboBox.SelectedItem }).Server
    $SCCMSiteCodeText.Text=($CountryNames | Where { $_.Country -eq $SCCMCountryComboBox.SelectedItem }).SiteCode
  })
  $SCCMGroupBox.Controls.Add($SCCMCountryComboBox)
 
  $SCCMServerLabel=New-Object System.Windows.Forms.Label
  $SCCMServerLabel.Text="SCCM Server:"
  $SCCMServerLabel.Location=New-Object System.Drawing.Size(15,55)
  $SCCMServerLabel.Size=New-Object System.Drawing.Size(80,20)
  $SCCMGroupBox.Controls.Add($SCCMServerLabel)
 
  $SCCMServerText=New-Object System.Windows.Forms.TextBox
  $SCCMServerText.Text=$SCCM_ConfigMgrSiteServer
  $SCCMServerText.Location=New-Object System.Drawing.Size(100,52)
  $SCCMServerText.Size=New-Object System.Drawing.Size(100,20)
  $SCCMGroupBox.Controls.Add($SCCMServerText)
 
  $SCCMSiteCodeLabel=New-Object System.Windows.Forms.Label
  $SCCMSiteCodeLabel.Text="Site Code:"
  $SCCMSiteCodeLabel.Location=New-Object System.Drawing.Size(15,85)
  $SCCMSiteCodeLabel.Size=New-Object System.Drawing.Size(80,20)
  $SCCMGroupBox.Controls.Add($SCCMSiteCodeLabel)
 
  $SCCMSiteCodeText=New-Object System.Windows.Forms.TextBox
  $SCCMSiteCodeText.Text=$SCCM_SiteCode
  $SCCMSiteCodeText.Location=New-Object System.Drawing.Size(100,85)
  $SCCMSiteCodeText.Size=New-Object System.Drawing.Size(100,20)
  $SCCMGroupBox.Controls.Add($SCCMSiteCodeText)

  $SCCMQueryGroupBox=New-Object System.Windows.Forms.GroupBox
  $SCCMQueryGroupBox.Location=New-Object System.Drawing.Size(15,145)
  $SCCMQueryGroupBox.Size=New-Object System.Drawing.Size(320,235)
  $SCCMQueryGroupBox.text="Scan - SCCM Query"
  $SCCMQueryGroupBox.BackColor="Transparent"
  $FormSCCMQuery.Controls.Add($SCCMQueryGroupBox)

  $SCCMQueryObjRadioButton=New-Object System.Windows.Forms.RadioButton
  $SCCMQueryObjRadioButton.Location=New-Object System.Drawing.Size(10,25)
  $SCCMQueryObjRadioButton.Size=New-Object System.Drawing.Size(210,20)
  $SCCMQueryObjRadioButton.Checked=$True
  $SCCMQueryObjRadioButton.Text="Query Objects Names"
  $SCCMQueryGroupBox.Controls.Add($SCCMQueryObjRadioButton)
  $SCCMQueryObjRadioButton.Add_Click({
    $SCCMQueryObjText.ReadOnly=-Not $SCCMQueryObjRadioButton.Checked
    $SCCMQueryIPText.ReadOnly=$SCCMQueryObjRadioButton.Checked
    $SCCMQueryManualText.ReadOnly=$SCCMQueryObjRadioButton.Checked
    $SCCMQueryObjText.Focus()
  })

  $SCCMQueryObjLabel=New-Object System.Windows.Forms.Label
  $SCCMQueryObjLabel.Text="Name pattern:"
  $SCCMQueryObjLabel.Location=New-Object System.Drawing.Size(15,50)
  $SCCMQueryObjLabel.Size=New-Object System.Drawing.Size(77,15)
  $SCCMQueryGroupBox.Controls.Add($SCCMQueryObjLabel)

  $SCCMQueryObjText=New-Object System.Windows.Forms.TextBox
  $SCCMQueryObjText.Text=""
  $SCCMQueryObjText.Location=New-Object System.Drawing.Size(100,47)
  $SCCMQueryObjText.Size=New-Object System.Drawing.Size(190,20)
  $SCCMQueryGroupBox.Controls.Add($SCCMQueryObjText)

  $SCCMQueryIPRadioButton=New-Object System.Windows.Forms.RadioButton
  $SCCMQueryIPRadioButton.Location=New-Object System.Drawing.Size(10,80)
  $SCCMQueryIPRadioButton.Size=New-Object System.Drawing.Size(210,20)
  $SCCMQueryIPRadioButton.Checked=$False
  $SCCMQueryIPRadioButton.Text="Query Objects IP"
  $SCCMQueryGroupBox.Controls.Add($SCCMQueryIPRadioButton)
  $SCCMQueryIPRadioButton.Add_Click({
    $SCCMQueryIPText.ReadOnly=-Not $SCCMQueryIPRadioButton.Checked
    $SCCMQueryObjText.ReadOnly=$SCCMQueryIPRadioButton.Checked
    $SCCMQueryManualText.ReadOnly=$SCCMQueryIPRadioButton.Checked
    $SCCMQueryIPText.Focus()
  })

  $SCCMQueryIPLabel=New-Object System.Windows.Forms.Label
  $SCCMQueryIPLabel.Text="IP pattern:"
  $SCCMQueryIPLabel.Location=New-Object System.Drawing.Size(15,105)
  $SCCMQueryIPLabel.Size=New-Object System.Drawing.Size(77,15)
  $SCCMQueryGroupBox.Controls.Add($SCCMQueryIPLabel)
 
  $SCCMQueryIPText=New-Object System.Windows.Forms.TextBox
  $SCCMQueryIPText.Text=""
  $SCCMQueryIPText.Location=New-Object System.Drawing.Size(100,102)
  $SCCMQueryIPText.Size=New-Object System.Drawing.Size(190,20)
  $SCCMQueryIPText.ReadOnly=$True
  $SCCMQueryGroupBox.Controls.Add($SCCMQueryIPText)

  $SCCMQueryManualRadioButton=New-Object System.Windows.Forms.RadioButton
  $SCCMQueryManualRadioButton.Location=New-Object System.Drawing.Size(10,135)
  $SCCMQueryManualRadioButton.Size=New-Object System.Drawing.Size(210,20)
  $SCCMQueryManualRadioButton.Checked=$False
  $SCCMQueryManualRadioButton.Text="Enter a Manual Query"
  $SCCMQueryGroupBox.Controls.Add($SCCMQueryManualRadioButton)
  $SCCMQueryManualRadioButton.Add_Click({
    $SCCMQueryManualText.ReadOnly=-Not $SCCMQueryManualRadioButton.Checked
    $SCCMQueryObjText.ReadOnly=$SCCMQueryManualRadioButton.Checked
    $SCCMQueryIPText.ReadOnly=$SCCMQueryManualRadioButton.Checked
    $SCCMQueryManualText.Focus() 
  })

  $SCCMQueryManualText=New-Object System.Windows.Forms.TextBox
  $SCCMQueryManualText.Text=""
  $SCCMQueryManualText.ReadOnly=$True
  $SCCMQueryManualText.Location=New-Object System.Drawing.Size(25,160)
  $SCCMQueryManualText.Size=New-Object System.Drawing.Size(270,60)
  $SCCMQueryManualText.ScrollBars="Vertical"
  $SCCMQueryManualText.Multiline=$True
  $SCCMQueryGroupBox.Controls.Add($SCCMQueryManualText)

  $SCCMQueryButton=New-Object System.Windows.Forms.Button
  $SCCMQueryButton.Location=New-Object System.Drawing.Size(70,395)
  $SCCMQueryButton.Size=New-Object System.Drawing.Size(80,40)
  $SCCMQueryButton.Text="Query"
  $SCCMQueryButton.BackColor='#FFCCCCCC'
  $SCCMQueryButton.Add_Click( {  # Save the country settings
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SCCM_ConfigMgrSiteServer" -Value $SCCMServerText.Text
    Set-ItemProperty -Path HKCU:\Software\Hydra\5 -Name "SCCM_SiteCode" -Value $SCCMSiteCodeText.Text
    $Script:SCCM_ConfigMgrSiteServer=$SCCMServerText.Text
    $Script:SCCM_SiteCode=$SCCMSiteCodeText.Text
    if ($SCCMQueryObjRadioButton.Checked) { Get-ObjectsSCCM "Object" } 
    if ($SCCMQueryIPRadioButton.Checked) { Get-ObjectsSCCM "IP" }
    if ($SCCMQueryManualRadioButton.Checked) { Get-ObjectsSCCM "Manual" }
  } )
  $FormSCCMQuery.Controls.Add($SCCMQueryButton)

  $SCCMCancelButton=New-Object System.Windows.Forms.Button
  $SCCMCancelButton.Location=New-Object System.Drawing.Size(200,395)
  $SCCMCancelButton.Size=New-Object System.Drawing.Size(80,40)
  $SCCMCancelButton.Text="Cancel"
  $SCCMCancelButton.BackColor='#FFCCCCCC'
  $SCCMCancelButton.Add_Click( { $FormSCCMQuery.Close() } )
  $FormSCCMQuery.Controls.Add($SCCMCancelButton)

  # AD query window

  $FormADQuery=New-Object Windows.Forms.Form
  $FormADQuery.Width=550
  $FormADQuery.Height=315
  $FormADQuery.Text="AD Query"
  $FormADQuery.BackColor=$ColorBackground
  $FormADQuery.Add_Load( { $FormADQuery.BackColor=$ColorBackground } )
  $FormADQuery.FormBorderStyle='FixedDialog'
  $FormADQuery.StartPosition="CenterParent"
  $FormADQuery.MinimizeBox=$False
  $FormADQuery.MaximizeBox=$False
  $FormADQuery.WindowState="Normal"
  $FormADQuery.KeyPreview=$True
  $FormADQuery.Add_KeyDown( {
    if ($_.KeyCode -eq "Escape") { $FormADQuery.Close() } 
  } )
  $FormADQuery.SizeGripStyle="Hide"

  $ADQueryGroupBox=New-Object System.Windows.Forms.GroupBox
  $ADQueryGroupBox.Location=New-Object System.Drawing.Size(15,15)
  $ADQueryGroupBox.Size=New-Object System.Drawing.Size(520,200)
  $ADQueryGroupBox.text="AD Queries"
  $ADQueryGroupBox.BackColor="Transparent"
  $FormADQuery.Controls.Add($ADQueryGroupBox)

  $ADQueryComboBox=New-Object System.Windows.Forms.ComboBox
  $ADQueryComboBox.Location=New-Object System.Drawing.Point(13, 30)
  $ADQueryComboBox.Size=New-Object System.Drawing.Size(230, 310)
  $ADQueryComboBox.DropDownStyle='DropDownList'
  $ADQueryComboBox.AutoCompleteMode='None'
  $ADQueryComboBox.AutoCompleteSource='ListItems'
  if (Test-Path $ADQueriesList) {  # Add the AD queries names in the combobox
    $ADQueryList=Import-Csv -Delimiter ";" -Path $ADQueriesList -Header QueryName, QueryDefinition
    foreach ($Query in $ADQueryList) {
     [void]$ADQueryComboBox.Items.Add($Query.QueryName)
    }
    $ADQueryComboBox.SelectedIndex=0
  }
  else {  # No AD queries file found
    $ADQueryComboBox.Enabled=$False
  }
  #$ADQueryComboBox.SelectedIndex=0
  $ADQueryComboBox_SelectedIndexChanged={  # Display the query definition depending on the query name selected
    if ($ADQueryComboBox.Text -notlike "*--*") { $ADQueryText.Text=$ADQueryList[$ADQueryComboBox.SelectedIndex].QueryDefinition }
  }  
  $ADQueryComboBox.Add_SelectedIndexChanged( $ADQueryComboBox_SelectedIndexChanged )
  $ADQueryGroupBox.Controls.Add($ADQueryComboBox)

  $ADQueryText=New-Object System.Windows.Forms.TextBox
  $ADQueryText.Text=""
  $ADQueryText.Location=New-Object System.Drawing.Size(50,77)
  $ADQueryText.Multiline=$True
  $ADQueryText.Size=New-Object System.Drawing.Size(430,60)
  $ADQueryGroupBox.Controls.Add($ADQueryText) 
  
  $ADQueryFilterLabel=New-Object System.Windows.Forms.Label
  $ADQueryFilterLabel.Text="Filter pattern:"
  $ADQueryFilterLabel.Location=New-Object System.Drawing.Size(15,160)
  $ADQueryFilterLabel.Size=New-Object System.Drawing.Size(75,15)
  $ADQueryGroupBox.Controls.Add($ADQueryFilterLabel)

  $ADQueryFilterText=New-Object System.Windows.Forms.TextBox
  $ADQueryFilterText.Text="*"
  $ADQueryFilterText.Location=New-Object System.Drawing.Size(100,155)
  $ADQueryFilterText.Size=New-Object System.Drawing.Size(80,20)
  $ADQueryGroupBox.Controls.Add($ADQueryFilterText)  

  $ADQueryButton=New-Object System.Windows.Forms.Button
  $ADQueryButton.Text="Query AD"
  $ADQueryButton.Enabled=$True
  $ADQueryButton.Location=New-Object System.Drawing.Size(170,230)
  $ADQueryButton.Size=New-Object System.Drawing.Size(80,40)
  $ADQueryButton.BackColor='#FFCCCCCC'
  $ADQueryButton.Add_Click( { 
    Get-ObjectsAD 
  } )
  $FormADQuery.Controls.Add($ADQueryButton)

  $ADCancelButton=New-Object System.Windows.Forms.Button
  $ADCancelButton.Text="Cancel"
  $ADCancelButton.Enabled=$True
  $ADCancelButton.Location=New-Object System.Drawing.Size(300,230)
  $ADCancelButton.Size=New-Object System.Drawing.Size(80,40)
  $ADCancelButton.BackColor='#FFCCCCCC'
  $ADCancelButton.Add_Click( { $FormADQuery.Close() } )
  $FormADQuery.Controls.Add($ADCancelButton)

  # IP range window

  $FormIPRange=New-Object Windows.Forms.Form
  $FormIPRange.Width=350
  $FormIPRange.Height=180
  $FormIPRange.Text="Create an IP range"
  $FormIPRange.BackColor=$ColorBackground
  $FormIPRange.Add_Load( { $FormIPRange.BackColor=$ColorBackground } )
  $FormIPRange.FormBorderStyle='FixedDialog'
  $FormIPRange.StartPosition="CenterParent"
  $FormIPRange.MinimizeBox=$False
  $FormIPRange.MaximizeBox=$False
  $FormIPRange.WindowState="Normal"
  $FormIPRange.KeyPreview=$True
  $FormIPRange.Add_KeyDown( {
    if ($_.KeyCode -eq "Escape") { $FormIPRange.Close() } 
  } )
  $FormIPRange.SizeGripStyle="Hide" 

  $IPRangeLabel=New-Object System.Windows.Forms.Label
  $IPRangeLabel.Text="Enter the IP range:"
  $IPRangeLabel.Location=New-Object System.Drawing.Size(15,15)
  $IPRangeLabel.Size=New-Object System.Drawing.Size(120,15)
  $FormIPRange.Controls.Add($IPRangeLabel)

  $IPRangeFromLabel=New-Object System.Windows.Forms.Label
  $IPRangeFromLabel.Text="From:"
  $IPRangeFromLabel.Location=New-Object System.Drawing.Size(30,50)
  $IPRangeFromLabel.Size=New-Object System.Drawing.Size(40,15)
  $FormIPRange.Controls.Add($IPRangeFromLabel)

  $IPRangeFromText=New-Object System.Windows.Forms.TextBox
  $IPRangeFromText.Text=""
  $IPRangeFromText.Location=New-Object System.Drawing.Size(70,47)
  $IPRangeFromText.Size=New-Object System.Drawing.Size(90,20)
  $FormIPRange.Controls.Add($IPRangeFromText)

  $IPRangeToLabel=New-Object System.Windows.Forms.Label
  $IPRangeToLabel.Text="To:"
  $IPRangeToLabel.Location=New-Object System.Drawing.Size(175,50)
  $IPRangeToLabel.Size=New-Object System.Drawing.Size(25,15)
  $FormIPRange.Controls.Add($IPRangeToLabel)

  $IPRangeToText=New-Object System.Windows.Forms.TextBox
  $IPRangeToText.Text=""
  $IPRangeToText.Location=New-Object System.Drawing.Size(200,47)
  $IPRangeToText.Size=New-Object System.Drawing.Size(90,20)
  $FormIPRange.Controls.Add($IPRangeToText)

  $IPRangeCreateButton=New-Object System.Windows.Forms.Button
  $IPRangeCreateButton.Text="Create"
  $IPRangeCreateButton.Enabled=$True
  $IPRangeCreateButton.Location=New-Object System.Drawing.Size(80,90)
  $IPRangeCreateButton.Size=New-Object System.Drawing.Size(70,30)
  $IPRangeCreateButton.BackColor='#FFCCCCCC'
  $IPRangeCreateButton.Add_Click( { 
    Get-IPRange 
  } )
  $FormIPRange.Controls.Add($IPRangeCreateButton)

  $IPRangeCancelButton=New-Object System.Windows.Forms.Button
  $IPRangeCancelButton.Text="Cancel"
  $IPRangeCancelButton.Enabled=$True
  $IPRangeCancelButton.Location=New-Object System.Drawing.Size(190,90)
  $IPRangeCancelButton.Size=New-Object System.Drawing.Size(70,30)
  $IPRangeCancelButton.BackColor='#FFCCCCCC'
  $IPRangeCancelButton.Add_Click( { $FormIPRange.Close() } )
  $FormIPRange.Controls.Add($IPRangeCancelButton)

  # About window

  $FormAbout=New-Object Windows.Forms.Form
  $FormAbout.Width=760
  $FormAbout.Height=362
  $FormAbout.BackColor='Black'
  $FormAbout.FormBorderStyle='None'
  $FormAbout.KeyPreview=$True
  $FormAbout.Add_KeyDown( {
    if ($_.KeyCode -eq "Escape") { $FormAbout.Close() } 
  } )
  $FormAbout.StartPosition="CenterParent"
  $FormAbout.Add_Load({
    $FormAboutPictureBox.Visible=$True
    $FormAboutHistoryRichTextBox.Visible=$False
  })

  $FormAboutPictureBox=New-Object Windows.Forms.PictureBox
  $FormAboutPictureBox.Width=510
  $FormAboutPictureBox.Height=360
  $FormAboutPictureBox.Top=1
  $FormAboutPictureBox.Left=1
  $FormAboutPictureBox.Image=$LogoImageFile
  $FormAbout.Controls.Add($FormAboutPictureBox)

  $FormAboutHistoryRichTextBox=New-Object System.Windows.Forms.RichTextBox
  $FormAboutHistoryRichTextBox.Location=New-Object System.Drawing.Size(1 ,1)
  $FormAboutHistoryRichTextBox.Size=New-Object System.Drawing.Size(510,360)
  $FormAboutHistoryRichTextBox.MultiLine=$True
  $FormAboutHistoryRichTextBox.Visible=$False
  $FormAboutHistoryRichTextBox.BackColor='LightGray'
  $FormAboutHistoryRichTextBox.ReadOnly=$True
  $FormAboutHistoryRichTextBox.WordWrap=$False
  $FormAboutHistoryRichTextBox.BorderStyle='None'
  if (Test-Path "$HydraDocsPath\Hydra_History.txt") {  # Load the history if it exists
    foreach ($Line in $((Get-Content "$HydraDocsPath\Hydra_History.txt") -split "`n`r")) { $FormAboutHistoryRichTextBox.SelectedText="$Line `n"}
  }
  $FormAbout.Controls.Add($FormAboutHistoryRichTextBox)

  $FormAboutRichTextBox=New-Object System.Windows.Forms.RichTextBox
  $FormAboutRichTextBox.Location=New-Object System.Drawing.Size(530, 15)
  $FormAboutRichTextBox.Size=New-Object System.Drawing.Size(200,250)
  $FormAboutRichTextBox.MultiLine=$True
  $FormAboutRichTextBox.Visible=$True
  $FormAboutRichTextBox.BackColor='Black'
  $FormAboutRichTextBox.ReadOnly=$True
  $FormAboutRichTextBox.WordWrap=$True
  $FormAboutRichTextBox.BorderStyle='None'
  $FormAboutRichTextBox.SelectionFont=New-Object Drawing.Font("Tahoma", 12, [Drawing.FontStyle]::Underline)
  $FormAboutRichTextBox.SelectionColor=[Drawing.Color]::Red
  $FormAboutRichTextBox.SelectedText="Hydra 5 (IronSnap)"
  $FormAboutRichTextBox.SelectionFont=New-Object Drawing.Font("Tahoma", 10)
  $FormAboutRichTextBox.SelectionColor=[Drawing.Color]::White
  $FormAboutRichTextBox.SelectedText="`n`n`nDeveloped by`n  Miguel Angel Torrecilla `n`n`n"
  $FormAboutRichTextBox.SelectionFont=New-Object Drawing.Font("Tahoma", 8)
  $FormAboutRichTextBox.SelectionColor=[Drawing.Color]::White
  $FormAboutRichTextBox.SelectedText="Credits:`n`nCode extensions, evolution & bug fixes`n  Miguel Angel Torrecilla`n`n"
  $FormAboutRichTextBox.SelectionFont=New-Object Drawing.Font("Tahoma", 8)
  $FormAboutRichTextBox.SelectionColor=[Drawing.Color]::White
  $FormAboutRichTextBox.SelectedText="About this software`n  This software and the commands contained are registered for HYDRA application which is also under EU intelectual property Law, the cost and license purchase are up to 5000€ this purchase allow to use distribute edit and reuse. No Refund`n`n"
  $FormAboutRichTextBox.SelectionFont=New-Object Drawing.Font("Tahoma", 8)
  $FormAboutRichTextBox.SelectionColor=[Drawing.Color]::White
  $FormAboutRichTextBox.SelectedText="Purchasing`n  It can be purchased by direct contact to MiguelAngelTorrecilla@outlook.com specify HYDRA purchase in subject"
  $FormAboutRichTextBox.Add_SelectionChanged( { $FormAboutRichTextBox.SelectionLength=0; $FormAboutCloseButton.Focus() })
  $FormAbout.Controls.Add($FormAboutRichTextBox)

  $FormAboutHistoryButton=New-Object System.Windows.Forms.Button
  $FormAboutHistoryButton.Text="History"
  $FormAboutHistoryButton.Enabled=$True
  $FormAboutHistoryButton.Visible=(Test-Path "$HydraDocsPath\Hydra_History.txt")
  $FormAboutHistoryButton.BackColor='Silver'
  $FormAboutHistoryButton.ForeColor='Black'
  $FormAboutHistoryButton.FlatStyle='Flat'
  $FormAboutHistoryButton.FlatAppearance.BorderSize=0
  $FormAboutHistoryButton.Location=New-Object System.Drawing.Size(550,300)
  $FormAboutHistoryButton.Size=New-Object System.Drawing.Size(60,25)
  $FormAbout.Controls.Add($FormAboutHistoryButton)
  $FormAboutHistoryButton.Add_Click( { 
    $FormAboutHistoryRichTextBox.Visible=!($FormAboutHistoryRichTextBox.Visible)
    $FormAboutPictureBox.Visible=!($FormAboutPictureBox.Visible)
  } )
  
  $FormAboutCloseButton=New-Object System.Windows.Forms.Button
  $FormAboutCloseButton.Text="Close"
  $FormAboutCloseButton.Enabled=$True
  $FormAboutCloseButton.BackColor='Silver'
  $FormAboutCloseButton.ForeColor='Black'
  $FormAboutCloseButton.FlatStyle='Flat'
  $FormAboutCloseButton.FlatAppearance.BorderSize=0
  $FormAboutCloseButton.Location=New-Object System.Drawing.Size(660,300)
  $FormAboutCloseButton.Size=New-Object System.Drawing.Size(60,25)
  $FormAboutCloseButton.Add_Click( { $FormAbout.Close() } )
  $FormAbout.Controls.Add($FormAboutCloseButton)
  $FormAboutCloseButton.Select()

  # Export window

  $FormExport=New-Object System.Windows.Forms.Form
  $FormExport.Size=New-Object System.Drawing.Size(336,280)
  $FormExport.FormBorderStyle="FixedDialog"
  $FormExport.Text="Export"
  $FormExport.StartPosition="CenterParent"
  $FormExport.BackColor=$ColorBackground
  $FormExport.Add_Load( { 
    $FormExport.BackColor=$ColorBackground 
    if ($ExportFormat -ne 3) {
      $FormExportExportButton.Text="To File..."
    }
    else {
      $FormExportExportButton.Text="Send Mail"
    }
  } )
  $FormExport.MinimizeBox=$False
  $FormExport.MaximizeBox=$False
  $FormExport.WindowState="Normal"
  $FormExport.SizeGripStyle="Hide"
  $FormExport.KeyPreview=$True
  $FormExport.Add_KeyDown( {
    if ($_.KeyCode -eq "Escape") { $FormExport.Close() } 
  } )
  $FormExport.Visible=$False

  $FormExportColumnsGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormExportColumnsGroupBox.Location=New-Object System.Drawing.Size(15,15)
  $FormExportColumnsGroupBox.Size=New-Object System.Drawing.Size(140,155)
  $FormExportColumnsGroupBox.Text="Columns"
  $FormExportColumnsGroupBox.BackColor="Transparent"
  $FormExport.Controls.Add($FormExportColumnsGroupBox)

  $FormExportLabelArray="Object","Task Result","Step","Sequence Name"
  $FormExportLabel=1..5
  $FormExportColCheckBox=1..5
  for ($i=1; $i -le 4; $i++) {
    $FormExportLabel[$i]=New-Object System.Windows.Forms.Label
    $FormExportLabel[$i].Text=$FormExportLabelArray[$i-1]
    $FormExportLabel[$i].Location=New-Object System.Drawing.Size(15, $(30+25*$($i-1)))
    $FormExportLabel[$i].Size=New-Object System.Drawing.Size(90,15)
    $FormExportColumnsGroupBox.Controls.Add($FormExportLabel[$i])
    $FormExportColCheckBox[$i]=New-Object System.Windows.Forms.Checkbox
    $FormExportColCheckBox[$i].Location=New-Object System.Drawing.Size(110, $(30+25*$($i-1)))
    $FormExportColCheckBox[$i].Size=New-Object System.Drawing.Size(20,20)
    $FormExportColCheckBox[$i].Checked=$True
    $FormExportColumnsGroupBox.Controls.Add($FormExportColCheckBox[$i])
  }

  $FormExportHeaderGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormExportHeaderGroupBox.Location=New-Object System.Drawing.Size(15,180)
  $FormExportHeaderGroupBox.Size=New-Object System.Drawing.Size(140,55)
  $FormExportHeaderGroupBox.Text="Header"
  $FormExportHeaderGroupBox.BackColor="Transparent"
  $FormExport.Controls.Add($FormExportHeaderGroupBox)

  $FormExportHeaderLabel=New-Object System.Windows.Forms.Label
  $FormExportHeaderLabel.Text="Export header"
  $FormExportHeaderLabel.Location=New-Object System.Drawing.Size(15, 25)
  $FormExportHeaderLabel.Size=New-Object System.Drawing.Size(90,15)
  $FormExportHeaderGroupBox.Controls.Add($FormExportHeaderLabel)

  $FormExportHeaderCheckBox=New-Object System.Windows.Forms.Checkbox
  $FormExportHeaderCheckBox.Location=New-Object System.Drawing.Size(110,25)
  $FormExportHeaderCheckBox.Size=New-Object System.Drawing.Size(20,20)
  $FormExportHeaderCheckBox.Checked=$True
  $FormExportHeaderGroupBox.Controls.Add($FormExportHeaderCheckBox)

  $FormExportSelectionGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormExportSelectionGroupBox.Location=New-Object System.Drawing.Size(175,15)
  $FormExportSelectionGroupBox.Size=New-Object System.Drawing.Size(140,90)
  $FormExportSelectionGroupBox.Text="Objects Selection"
  $FormExportSelectionGroupBox.BackColor="Transparent"
  $FormExport.Controls.Add($FormExportSelectionGroupBox)

  $FormExportAllLabel=New-Object System.Windows.Forms.Label
  $FormExportAllLabel.Text="Export all"
  $FormExportAllLabel.Location=New-Object System.Drawing.Size(15, 30)
  $FormExportAllLabel.Size=New-Object System.Drawing.Size(90,15)
  $FormExportSelectionGroupBox.Controls.Add($FormExportAllLabel)

  $FormExportAllCheckBox=New-Object System.Windows.Forms.Checkbox
  $FormExportAllCheckBox.Location=New-Object System.Drawing.Size(110,30)
  $FormExportAllCheckBox.Size=New-Object System.Drawing.Size(20,20)
  $FormExportAllCheckBox.Checked=$True
  $FormExportAllCheckBox.Add_Click( { $FormExportSelectionCheckBox.Checked=-Not $FormExportAllCheckBox.Checked } )
  $FormExportSelectionGroupBox.Controls.Add($FormExportAllCheckBox)

  $FormExportSelectionLabel=New-Object System.Windows.Forms.Label
  $FormExportSelectionLabel.Text="Export selection"
  $FormExportSelectionLabel.Location=New-Object System.Drawing.Size(15, 55)
  $FormExportSelectionLabel.Size=New-Object System.Drawing.Size(90,15)
  $FormExportSelectionGroupBox.Controls.Add($FormExportSelectionLabel)
    
  $FormExportSelectionCheckBox=New-Object System.Windows.Forms.Checkbox
  $FormExportSelectionCheckBox.Location=New-Object System.Drawing.Size(110,55)
  $FormExportSelectionCheckBox.Size=New-Object System.Drawing.Size(20,20)
  $FormExportSelectionCheckBox.Checked=$False
  $FormExportSelectionCheckBox.Add_Click( { $FormExportAllCheckBox.Checked=-Not $FormExportSelectionCheckBox.Checked } )
  $FormExportSelectionGroupBox.Controls.Add($FormExportSelectionCheckBox)

  $FormExportColorsGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormExportColorsGroupBox.Location=New-Object System.Drawing.Size(175,115)
  $FormExportColorsGroupBox.Size=New-Object System.Drawing.Size(140,55)
  $FormExportColorsGroupBox.Text="Colors"
  $FormExportColorsGroupBox.BackColor="Transparent"
  $FormExport.Controls.Add($FormExportColorsGroupBox)

  $FormExportColorLabel=New-Object System.Windows.Forms.Label
  $FormExportColorLabel.Text="Export colors"
  $FormExportColorLabel.Location=New-Object System.Drawing.Size(15, 25)
  $FormExportColorLabel.Size=New-Object System.Drawing.Size(80,15)
  $FormExportColorsGroupBox.Controls.Add($FormExportColorLabel)

  $FormExportColorCheckBox=New-Object System.Windows.Forms.Checkbox
  $FormExportColorCheckBox.Location=New-Object System.Drawing.Size(110,25)
  $FormExportColorCheckBox.Size=New-Object System.Drawing.Size(20,20)
  $FormExportColorCheckBox.Checked=$True
  $FormExportColorsGroupBox.Controls.Add($FormExportColorCheckBox)

  $FormExportCancelButton=New-Object System.Windows.Forms.Button
  $FormExportCancelButton.Text="Cancel"
  $FormExportCancelButton.Enabled=$True
  $FormExportCancelButton.Location=New-Object System.Drawing.Size(170,205)
  $FormExportCancelButton.Size=New-Object System.Drawing.Size(70,30)
  $FormExportCancelButton.BackColor='#FFCCCCCC'
  $FormExportCancelButton.Add_Click( { $FormExport.Close() } )
  $FormExport.Controls.Add($FormExportCancelButton)

  $FormExportExportButton=New-Object System.Windows.Forms.Button
  $FormExportExportButton.Text="To File..."
  $FormExportExportButton.Enabled=$True
  $FormExportExportButton.Location=New-Object System.Drawing.Size(250,205)
  $FormExportExportButton.Size=New-Object System.Drawing.Size(70,30)
  $FormExportExportButton.BackColor='#FFCCCCCC'
  $FormExportExportButton.Add_Click( { Export-Result $ExportFormat } )
  $FormExport.Controls.Add($FormExportExportButton)

  # Settings windows

  $FormSettings=New-Object System.Windows.Forms.Form
  $FormSettings.Size='470, 380'
  $FormSettings.StartPosition='CenterParent'
  $FormSettings.FormBorderStyle="FixedDialog"
  $FormSettings.BackColor=$ColorBackground
  $FormSettings.MinimizeBox=$False
  $FormSettings.MaximizeBox=$False
  $FormSettings.WindowState="Normal"
  $FormSettings.Text="Settings"
  $FormSettings.KeyPreview=$True
  $FormSettings.Add_KeyDown( {
    if ($_.KeyCode -eq "Escape") { $FormSettings.Close() } 
  } )
  $FormSettings.Add_Load( { $FormSettings.BackColor=$ColorBackground; Save-CurrentSettings } )
  $FormSettings.Add_Shown( { $FormSettingsOKButton.Focus() } )
  $FormSettings.SizeGripStyle="Hide"

  $FormSettingsMenuToolStrip=New-Object System.Windows.Forms.ToolStrip
  $FormSettingsMenuToolStrip.AutoSize=$False
  $FormSettingsMenuToolStrip.Height=70
  $FormSettingsMenuToolStrip.Location="0,0"
  $FormSettingsMenuToolStrip.ImageScalingSize="64,64"
  $FormSettingsMenuToolStrip.RenderMode="System"
  $FormSettingsMenuToolStrip.GripStyle="Hidden"
  $FormSettingsMenuToolStrip.BackColor=$ColorBackground
  [void]$FormSettings.Controls.Add($FormSettingsMenuToolStrip)

  Add-ToolStripButton "Paths" $Icon01 { Set-SettingsSubMenu "Paths" } $FormSettingsMenuToolStrip 68
  Add-ToolStripButton "Colors" $Icon02 { Set-SettingsSubMenu "Colors" } $FormSettingsMenuToolStrip 68
  Add-ToolStripButton "Misc" $Icon03 { Set-SettingsSubMenu "Misc" } $FormSettingsMenuToolStrip 68
  Add-ToolStripButton "Mail" $Icon04 { Set-SettingsSubMenu "Mail" } $FormSettingsMenuToolStrip 68
  Add-ToolStripButton "Groups" $Icon05 { Set-SettingsSubMenu "Groups" } $FormSettingsMenuToolStrip 68
  Add-ToolStripButton "Tabs" $Icon06 { Set-SettingsSubMenu "Tabs" } $FormSettingsMenuToolStrip 68

  $FormSettingsPathsGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormSettingsPathsGroupBox.Location=New-Object System.Drawing.Size(15,90)
  $FormSettingsPathsGroupBox.Size=New-Object System.Drawing.Size(380,190)
  $FormSettingsPathsGroupBox.Text="Paths"
  $FormSettingsPathsGroupBox.Visible=$True
  $FormSettingsPathsGroupBox.BackColor="Transparent"
  $FormSettings.Controls.Add($FormSettingsPathsGroupBox)

  $FormSettingsPathsLabelText=@("CSV Temp File";"XLSX Temp File";"HTML Temp File";"Hydra Log File";"Central Log Path")
  $FormSettingsPathsValue=@($CSVTempPath;$XLSXTempPath;$HTMLTempPath;$LogFilePath;$CentralLogPath)
  $FormSettingsPathsVariable=@("CSVTempPath";"XLSXTempPath";"HTMLTempPath";"LogFilePath";"CentralLogPath")
  $FormSettingsPathsLabel=0..4
  $FormSettingsPathsText=0..4
  
  for ($i=0; $i -le 4; $i++) {
    $FormSettingsPathsLabel[$i]=New-Object System.Windows.Forms.Label
    $FormSettingsPathsLabel[$i].Text=$FormSettingsPathsLabelText[$i]
    $FormSettingsPathsLabel[$i].Location=New-Object System.Drawing.Size(15, $(33+30*$i-1))
    $FormSettingsPathsLabel[$i].Size=New-Object System.Drawing.Size(100,15)
    $FormSettingsPathsGroupBox.Controls.Add($FormSettingsPathsLabel[$i])
    $FormSettingsPathsText[$i]=New-Object System.Windows.Forms.TextBox
    $FormSettingsPathsText[$i].Text=$FormSettingsPathsValue[$i]
    $FormSettingsPathsText[$i].Location=New-Object System.Drawing.Size(120, $(30+30*$i-1))
    $FormSettingsPathsText[$i].Size=New-Object System.Drawing.Size(240,20)
    $FormSettingsPathsGroupBox.Controls.Add($FormSettingsPathsText[$i])
  }

  $FormSettingsPathsLabel[3].Size=New-Object System.Drawing.Size(80,15)
  $FormSettingsLogCheckBox=New-Object System.Windows.Forms.Checkbox
  $FormSettingsLogCheckBox.Location=New-Object System.Drawing.Size(100,120)
  $FormSettingsLogCheckBox.Size=New-Object System.Drawing.Size(20,20)
  if ($LogFileEnabled -eq "True") { $FormSettingsLogCheckBox.Checked=$True }
  else { $FormSettingsLogCheckBox.Checked=$False }
  $FormSettingsPathsText[3].Enabled=$FormSettingsLogCheckBox.Checked
  $FormSettingsLogCheckBox.Add_Click( {
    $FormSettingsPathsText[3].Enabled=$FormSettingsLogCheckBox.Checked  
  } )
  $FormSettingsPathsGroupBox.Controls.Add($FormSettingsLogCheckBox)

  $FormSettingsColorsGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormSettingsColorsGroupBox.Location=New-Object System.Drawing.Size(15,90)
  $FormSettingsColorsGroupBox.Size=New-Object System.Drawing.Size(200,160)
  $FormSettingsColorsGroupBox.Text="Colors"
  $FormSettingsColorsGroupBox.BackColor="Transparent"
  $FormSettingsColorsGroupBox.Visible=$False
  $FormSettings.Controls.Add($FormSettingsColorsGroupBox)

  $FormSettingsColorsLabelText=@("OK";"BREAK";"STOP";"CANCELLED")
  $FormSettingsColorsLabel=0..4
  $FormSettingsColorsButton=0..4
  
   for ($i=0; $i -le 3; $i++) {
    $FormSettingsColorsLabel[$i]=New-Object System.Windows.Forms.Label
    $FormSettingsColorsLabel[$i].Text=$FormSettingsColorsLabelText[$i]
    $FormSettingsColorsLabel[$i].Location=New-Object System.Drawing.Size(15, $(30+30*$i-1))
    $FormSettingsColorsLabel[$i].Size=New-Object System.Drawing.Size(80,15)
    $FormSettingsColorsGroupBox.Controls.Add($FormSettingsColorsLabel[$i])

    $FormSettingsColorsButton[$i]=New-Object System.Windows.Forms.Button
    $FormSettingsColorsButton[$i].Location=New-Object System.Drawing.Size(95, $(26+30*$i-1))
    $FormSettingsColorsButton[$i].Size=New-Object System.Drawing.Size(90,25)
    $FormSettingsColorsButton[$i].Name=$FormSettingsColorsLabelText[$i]
    $FormSettingsColorsButton[$i].BackColor=$Colors.Get_Item($FormSettingsColorsLabelText[$i])
    $FormSettingsColorsButton[$i].Add_Click( { 
      $ColorResult=Show-PickColor $This.Name
      if ($ColorResult -ne "Cancel") { $This.BackColor=$ColorResult }
    } )
    $FormSettingsColorsGroupBox.Controls.Add($FormSettingsColorsButton[$i])
  }
  
  $FormSettingsColorsGUIGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormSettingsColorsGUIGroupBox.Location=New-Object System.Drawing.Size(230,90)
  $FormSettingsColorsGUIGroupBox.Size=New-Object System.Drawing.Size(200,130)
  $FormSettingsColorsGUIGroupBox.Text="Colors GUI"
  $FormSettingsColorsGUIGroupBox.BackColor="Transparent"
  $FormSettingsColorsGUIGroupBox.Visible=$False
  $FormSettings.Controls.Add($FormSettingsColorsGUIGroupBox)

  $FormSettingsColorsGUIBackLabel=New-Object System.Windows.Forms.Label
  $FormSettingsColorsGUIBackLabel.Text="Background"
  $FormSettingsColorsGUIBackLabel.Location=New-Object System.Drawing.Size(15, 30)
  $FormSettingsColorsGUIBackLabel.Size=New-Object System.Drawing.Size(80,15)
  $FormSettingsColorsGUIGroupBox.Controls.Add($FormSettingsColorsGUIBackLabel)

  $FormSettingsColorsGUIBackButton=New-Object System.Windows.Forms.Button
  $FormSettingsColorsGUIBackButton.Location=New-Object System.Drawing.Size(95, 26)
  $FormSettingsColorsGUIBackButton.Size=New-Object System.Drawing.Size(90,25)
  $FormSettingsColorsGUIBackButton.BackColor=$ColorBackground
  $FormSettingsColorsGUIBackButton.Add_Click( { 
    $ColorResult=Show-PickColor $This.Name
    if ($ColorResult -ne "Cancel") { $This.BackColor=$ColorResult }
  } )
  $FormSettingsColorsGUIGroupBox.Controls.Add($FormSettingsColorsGUIBackButton)

  $FormSettingsColorsGUISeqLabel=New-Object System.Windows.Forms.Label
  $FormSettingsColorsGUISeqLabel.Text="Sequences"
  $FormSettingsColorsGUISeqLabel.Location=New-Object System.Drawing.Size(15, 60)
  $FormSettingsColorsGUISeqLabel.Size=New-Object System.Drawing.Size(80,15)
  $FormSettingsColorsGUIGroupBox.Controls.Add($FormSettingsColorsGUISeqLabel)

  $FormSettingsColorsGUISeqButton=New-Object System.Windows.Forms.Button
  $FormSettingsColorsGUISeqButton.Location=New-Object System.Drawing.Size(95, 56)
  $FormSettingsColorsGUISeqButton.Size=New-Object System.Drawing.Size(90,25)
  $FormSettingsColorsGUISeqButton.BackColor=$ColorSequences
  $FormSettingsColorsGUISeqButton.Add_Click( { 
    $ColorResult=Show-PickColor $This.Name
    if ($ColorResult -ne "Cancel") { $This.BackColor=$ColorResult }
  } )
  $FormSettingsColorsGUIGroupBox.Controls.Add($FormSettingsColorsGUISeqButton)

  $FormSettingsColorsGUISeqRunLabel=New-Object System.Windows.Forms.Label
  $FormSettingsColorsGUISeqRunLabel.Text="Running Seq."
  $FormSettingsColorsGUISeqRunLabel.Location=New-Object System.Drawing.Size(15, 90)
  $FormSettingsColorsGUISeqRunLabel.Size=New-Object System.Drawing.Size(80,15)
  $FormSettingsColorsGUIGroupBox.Controls.Add($FormSettingsColorsGUISeqRunLabel)

  $FormSettingsColorsGUISeqRunButton=New-Object System.Windows.Forms.Button
  $FormSettingsColorsGUISeqRunButton.Location=New-Object System.Drawing.Size(95, 86)
  $FormSettingsColorsGUISeqRunButton.Size=New-Object System.Drawing.Size(90,25)
  $FormSettingsColorsGUISeqRunButton.BackColor=$ColorSequencesRunning
  $FormSettingsColorsGUISeqRunButton.Add_Click( { 
    $ColorResult=Show-PickColor $This.Name
    if ($ColorResult -ne "Cancel") { $This.BackColor=$ColorResult }
  } )
  $FormSettingsColorsGUIGroupBox.Controls.Add($FormSettingsColorsGUISeqRunButton)

  $FormSettingsSequenceSearchGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormSettingsSequenceSearchGroupBox.Location=New-Object System.Drawing.Size(15,90)
  $FormSettingsSequenceSearchGroupBox.Size=New-Object System.Drawing.Size(220,60)
  $FormSettingsSequenceSearchGroupBox.Text="Sequences Search Box"
  $FormSettingsSequenceSearchGroupBox.BackColor="Transparent"
  $FormSettingsSequenceSearchGroupBox.Visible=$False
  $FormSettings.Controls.Add($FormSettingsSequenceSearchGroupBox)

  $FormSettingsSequenceShowSearchRadioButton=New-Object System.Windows.Forms.RadioButton
  $FormSettingsSequenceShowSearchRadioButton.Location=New-Object System.Drawing.Size(15,25)
  $FormSettingsSequenceShowSearchRadioButton.Size=New-Object System.Drawing.Size(100,20)
  $FormSettingsSequenceShowSearchRadioButton.Checked=($ShowSearchBox -eq "True")
  $FormSettingsSequenceShowSearchRadioButton.Text="Show the Box"
  $FormSettingsSequenceSearchGroupBox.Controls.Add($FormSettingsSequenceShowSearchRadioButton)

  $FormSettingsSequenceShowHideRadioButton=New-Object System.Windows.Forms.RadioButton
  $FormSettingsSequenceShowHideRadioButton.Location=New-Object System.Drawing.Size(120,25)
  $FormSettingsSequenceShowHideRadioButton.Size=New-Object System.Drawing.Size(100,20)
  $FormSettingsSequenceShowHideRadioButton.Checked=($ShowSearchBox -ne "True")
  $FormSettingsSequenceShowHideRadioButton.Text="Hide the Box"
  $FormSettingsSequenceSearchGroupBox.Controls.Add($FormSettingsSequenceShowHideRadioButton)

  $FormSettingsMiscGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormSettingsMiscGroupBox.Location=New-Object System.Drawing.Size(15,230)
  $FormSettingsMiscGroupBox.Size=New-Object System.Drawing.Size(220,60)
  $FormSettingsMiscGroupBox.Text="Misc"
  $FormSettingsMiscGroupBox.BackColor="Transparent"
  $FormSettingsMiscGroupBox.Visible=$False
  $FormSettings.Controls.Add($FormSettingsMiscGroupBox)

  $FormSettingsSplashScreenCheckBox=New-Object System.Windows.Forms.Checkbox
  $FormSettingsSplashScreenCheckBox.Location=New-Object System.Drawing.Size(15,23)
  if ($NoSplashScreen -eq "True") { $FormSettingsSplashScreenCheckBox.Checked=$False }
  else { $FormSettingsSplashScreenCheckBox.Checked=$True }
  $FormSettingsSplashScreenCheckBox.Text="Splash Screen"
  $FormSettingsMiscGroupBox.Controls.Add($FormSettingsSplashScreenCheckBox)

  $FormSettingsDebugScreenCheckBox=New-Object System.Windows.Forms.Checkbox
  $FormSettingsDebugScreenCheckBox.Location=New-Object System.Drawing.Size(120,23)
  if ($DebugMode -eq 0) { $FormSettingsDebugScreenCheckBox.Checked=$False }
  else { $FormSettingsDebugScreenCheckBox.Checked=$True }
  $FormSettingsDebugScreenCheckBox.Text="Debug Mode"
  $FormSettingsMiscGroupBox.Controls.Add($FormSettingsDebugScreenCheckBox)

  $FormSettingsSequenceGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormSettingsSequenceGroupBox.Location=New-Object System.Drawing.Size(15,160)
  $FormSettingsSequenceGroupBox.Size=New-Object System.Drawing.Size(220,60)
  $FormSettingsSequenceGroupBox.Text="Sequences List"
  $FormSettingsSequenceGroupBox.BackColor="Transparent"
  $FormSettingsSequenceGroupBox.Visible=$False
  $FormSettings.Controls.Add($FormSettingsSequenceGroupBox)

  $FormSettingsSequenceExpandedRadioButton=New-Object System.Windows.Forms.RadioButton
  $FormSettingsSequenceExpandedRadioButton.Location=New-Object System.Drawing.Size(15,25)
  $FormSettingsSequenceExpandedRadioButton.Size=New-Object System.Drawing.Size(100,20)
  $FormSettingsSequenceExpandedRadioButton.Checked=($SequenceListExpanded -eq "True")
  $FormSettingsSequenceExpandedRadioButton.Text="List expanded"
  $FormSettingsSequenceGroupBox.Controls.Add($FormSettingsSequenceExpandedRadioButton)

  $FormSettingsSequenceCollapsedRadioButton=New-Object System.Windows.Forms.RadioButton
  $FormSettingsSequenceCollapsedRadioButton.Location=New-Object System.Drawing.Size(120,25)
  $FormSettingsSequenceCollapsedRadioButton.Size=New-Object System.Drawing.Size(100,20)
  $FormSettingsSequenceCollapsedRadioButton.Checked=($SequenceListExpanded -ne "True")
  $FormSettingsSequenceCollapsedRadioButton.Text="List collapsed"
  $FormSettingsSequenceGroupBox.Controls.Add($FormSettingsSequenceCollapsedRadioButton)

  $FormSettingsRowHeaderGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormSettingsRowHeaderGroupBox.Location=New-Object System.Drawing.Size(250,90)
  $FormSettingsRowHeaderGroupBox.Size=New-Object System.Drawing.Size(200,60)
  $FormSettingsRowHeaderGroupBox.Text="Rows Header"
  $FormSettingsRowHeaderGroupBox.BackColor="Transparent"
  $FormSettingsRowHeaderGroupBox.Visible=$False
  $FormSettings.Controls.Add($FormSettingsRowHeaderGroupBox)

  $FormSettingsRowHeaderGroupBoxHiddenRadioButton=New-Object System.Windows.Forms.RadioButton
  $FormSettingsRowHeaderGroupBoxHiddenRadioButton.Location=New-Object System.Drawing.Size(15,25)
  $FormSettingsRowHeaderGroupBoxHiddenRadioButton.Size=New-Object System.Drawing.Size(60,20)
  $FormSettingsRowHeaderGroupBoxHiddenRadioButton.Checked=($RowHeaderVisible -ne "True")
  $FormSettingsRowHeaderGroupBoxHiddenRadioButton.Text="Hidden"
  $FormSettingsRowHeaderGroupBox.Controls.Add($FormSettingsRowHeaderGroupBoxHiddenRadioButton)

  $FormSettingsRowHeaderGroupBoxVisibleRadioButton=New-Object System.Windows.Forms.RadioButton
  $FormSettingsRowHeaderGroupBoxVisibleRadioButton.Location=New-Object System.Drawing.Size(110,25)
  $FormSettingsRowHeaderGroupBoxVisibleRadioButton.Size=New-Object System.Drawing.Size(60,20)
  $FormSettingsRowHeaderGroupBoxVisibleRadioButton.Checked=($RowHeaderVisible -eq "True")
  $FormSettingsRowHeaderGroupBoxVisibleRadioButton.Text="Visible"
  $FormSettingsRowHeaderGroupBox.Controls.Add($FormSettingsRowHeaderGroupBoxVisibleRadioButton)

  $FormSettingsCheckBoxesGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormSettingsCheckBoxesGroupBox.Location=New-Object System.Drawing.Size(250,160)
  $FormSettingsCheckBoxesGroupBox.Size=New-Object System.Drawing.Size(200,60)
  $FormSettingsCheckBoxesGroupBox.Text="Checkboxes"
  $FormSettingsCheckBoxesGroupBox.BackColor="Transparent"
  $FormSettingsCheckBoxesGroupBox.Visible=$False
  $FormSettings.Controls.Add($FormSettingsCheckBoxesGroupBox)

  $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton=New-Object System.Windows.Forms.RadioButton
  $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Location=New-Object System.Drawing.Size(110,25)
  $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Size=New-Object System.Drawing.Size(90,20)
  $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Checked=($CheckBoxesKeepState -eq "True")
  $FormSettingsCheckBoxesGroupBoxKeepStateRadioButton.Text="Keep State"
  $FormSettingsCheckBoxesGroupBox.Controls.Add($FormSettingsCheckBoxesGroupBoxKeepStateRadioButton)

  $FormSettingsCheckBoxesGroupBoxResetStateRadioButton=New-Object System.Windows.Forms.RadioButton
  $FormSettingsCheckBoxesGroupBoxResetStateRadioButton.Location=New-Object System.Drawing.Size(15,25)
  $FormSettingsCheckBoxesGroupBoxResetStateRadioButton.Size=New-Object System.Drawing.Size(90,20)
  $FormSettingsCheckBoxesGroupBoxResetStateRadioButton.Checked=($CheckBoxesKeepState -ne "True")
  $FormSettingsCheckBoxesGroupBoxResetStateRadioButton.Text="Reset State"
  $FormSettingsCheckBoxesGroupBox.Controls.Add($FormSettingsCheckBoxesGroupBoxResetStateRadioButton)

  $FormSettingsMailGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormSettingsMailGroupBox.Location=New-Object System.Drawing.Size(15,90)
  $FormSettingsMailGroupBox.Size=New-Object System.Drawing.Size(380,160)
  $FormSettingsMailGroupBox.Text="Mail"
  $FormSettingsMailGroupBox.Visible=$False
  $FormSettingsMailGroupBox.BackColor="Transparent"
  $FormSettings.Controls.Add($FormSettingsMailGroupBox)

  $FormSettingsMailLabelText=@("SMTP Server";"Mail From";"Send To";"Reply To")
  $FormSettingsMailValue=@($EMailSMTPServer;$EMailSendFrom;$EMailSendTo;$EMailReplyTo)
  $FormSettingsMailVariable=@("EMailSMTPServer";"EMailSendFrom";"EMailSendTo";"EMailReplyTo")
  $FormSettingsMailLabel=0..3
  $FormSettingsMailText=0..3
  
  for ($i=0; $i -le 3; $i++) {
    $FormSettingsMailLabel[$i]=New-Object System.Windows.Forms.Label
    $FormSettingsMailLabel[$i].Text=$FormSettingsMailLabelText[$i]
    $FormSettingsMailLabel[$i].Location=New-Object System.Drawing.Size(15, $(30+30*$i-1))
    $FormSettingsMailLabel[$i].Size=New-Object System.Drawing.Size(100,15)
    $FormSettingsMailGroupBox.Controls.Add($FormSettingsMailLabel[$i])
    $FormSettingsMailText[$i]=New-Object System.Windows.Forms.TextBox
    $FormSettingsMailText[$i].Text=$FormSettingsMailValue[$i]
    $FormSettingsMailText[$i].Location=New-Object System.Drawing.Size(120, $(30+30*$i-1))
    $FormSettingsMailText[$i].Size=New-Object System.Drawing.Size(240,20)
    $FormSettingsMailGroupBox.Controls.Add($FormSettingsMailText[$i])
  }

  $FormSettingsGroupsWarningGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormSettingsGroupsWarningGroupBox.Location=New-Object System.Drawing.Size(15,90)
  $FormSettingsGroupsWarningGroupBox.Size=New-Object System.Drawing.Size(220,60)
  $FormSettingsGroupsWarningGroupBox.Text="Checkboxes on Warning"
  $FormSettingsGroupsWarningGroupBox.BackColor="Transparent"
  $FormSettingsGroupsWarningGroupBox.Visible=$False
  $FormSettings.Controls.Add($FormSettingsGroupsWarningGroupBox)

  $FormSettingsGroupsWarningUncheckedRadioButton=New-Object System.Windows.Forms.RadioButton
  $FormSettingsGroupsWarningUncheckedRadioButton.Location=New-Object System.Drawing.Size(15,25)
  $FormSettingsGroupsWarningUncheckedRadioButton.Size=New-Object System.Drawing.Size(100,20)
  $FormSettingsGroupsWarningUncheckedRadioButton.Checked=($GrpCheckedOnWarning -ne "True")
  $FormSettingsGroupsWarningUncheckedRadioButton.Text="De-selected"
  $FormSettingsGroupsWarningGroupBox.Controls.Add($FormSettingsGroupsWarningUncheckedRadioButton)

  $FormSettingsGroupsWarningCheckedRadioButton=New-Object System.Windows.Forms.RadioButton
  $FormSettingsGroupsWarningCheckedRadioButton.Location=New-Object System.Drawing.Size(120,25)
  $FormSettingsGroupsWarningCheckedRadioButton.Size=New-Object System.Drawing.Size(100,20)
  $FormSettingsGroupsWarningCheckedRadioButton.Checked=($GrpCheckedOnWarning -eq "True")
  $FormSettingsGroupsWarningCheckedRadioButton.Text="Selected"
  $FormSettingsGroupsWarningGroupBox.Controls.Add($FormSettingsGroupsWarningCheckedRadioButton)

  $FormSettingsGroupsThreadsGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormSettingsGroupsThreadsGroupBox.Location=New-Object System.Drawing.Size(15,160)
  $FormSettingsGroupsThreadsGroupBox.Size=New-Object System.Drawing.Size(220,60)
  $FormSettingsGroupsThreadsGroupBox.Text="Max. Threads in Cells"
  $FormSettingsGroupsThreadsGroupBox.BackColor="Transparent"
  $FormSettingsGroupsThreadsGroupBox.Visible=$False
  $FormSettings.Controls.Add($FormSettingsGroupsThreadsGroupBox)

  $FormSettingsGroupsThreadsVisibleRadioButton=New-Object System.Windows.Forms.RadioButton
  $FormSettingsGroupsThreadsVisibleRadioButton.Location=New-Object System.Drawing.Size(15,25)
  $FormSettingsGroupsThreadsVisibleRadioButton.Size=New-Object System.Drawing.Size(100,20)
  $FormSettingsGroupsThreadsVisibleRadioButton.Checked=($GrpShowThreads -eq "True")
  $FormSettingsGroupsThreadsVisibleRadioButton.Text="Show"
  $FormSettingsGroupsThreadsGroupBox.Controls.Add($FormSettingsGroupsThreadsVisibleRadioButton)

  $FormSettingsGroupsThreadsInvisibleRadioButton=New-Object System.Windows.Forms.RadioButton
  $FormSettingsGroupsThreadsInvisibleRadioButton.Location=New-Object System.Drawing.Size(120,25)
  $FormSettingsGroupsThreadsInvisibleRadioButton.Size=New-Object System.Drawing.Size(100,20)
  $FormSettingsGroupsThreadsInvisibleRadioButton.Checked=($GrpShowThreads -ne "True")
  $FormSettingsGroupsThreadsInvisibleRadioButton.Text="Hide"
  $FormSettingsGroupsThreadsGroupBox.Controls.Add($FormSettingsGroupsThreadsInvisibleRadioButton)

  $FormSettingsTabsLookGroupBox=New-Object System.Windows.Forms.GroupBox
  $FormSettingsTabsLookGroupBox.Location=New-Object System.Drawing.Size(15,90)
  $FormSettingsTabsLookGroupBox.Size=New-Object System.Drawing.Size(350,150)
  $FormSettingsTabsLookGroupBox.Text="Tabs Style"
  $FormSettingsTabsLookGroupBox.BackColor="Transparent"
  $FormSettingsTabsLookGroupBox.Visible=$False
  $FormSettings.Controls.Add($FormSettingsTabsLookGroupBox)

  $FormSettingsTabsLookCheckedRadioButton=0..2

  for ($i=0; $i -le 2; $i++) {
    $FormSettingsTabsLookCheckedRadioButton[$i]=New-Object System.Windows.Forms.RadioButton
    $FormSettingsTabsLookCheckedRadioButton[$i].Location=New-Object System.Drawing.Size(25, $(30+40*$i-1))
    $FormSettingsTabsLookCheckedRadioButton[$i].Size=New-Object System.Drawing.Size(20,20)
    $FormSettingsTabsLookCheckedRadioButton[$i].Checked=($i -eq $TabLook)
    $FormSettingsTabsLookGroupBox.Controls.Add($FormSettingsTabsLookCheckedRadioButton[$i])
    $TabPictureBox=New-Object System.Windows.Forms.PictureBox
    $TabPictureBox.Size=New-Object System.Drawing.Size(273,29)
    $TabPictureBox.Location=New-Object System.Drawing.Size(50,$(20+40*$i-1))
    $TabPictureBox.Image=$TabImage[$i]
    $FormSettingsTabsLookGroupBox.Controls.Add($TabPictureBox)
  }
  Set-TabStyle

  $FormSettingsOKButton=New-Object System.Windows.Forms.Button
  $FormSettingsOKButton.Location=New-Object System.Drawing.Size(15,315) 
  $FormSettingsOKButton.Size=New-Object System.Drawing.Size(80,25)
  $FormSettingsOKButton.Text="OK"
  $FormSettingsOKButton.BackColor="#FFCCCCCC"
  $FormSettingsOKButton.Add_Click( { 
    Set-UserSettings
    $FormSettings.Close()
    
  } )
  $FormSettings.Controls.Add($FormSettingsOKButton)

  $FormSettingsCancelButton=New-Object System.Windows.Forms.Button
  $FormSettingsCancelButton.Location=New-Object System.Drawing.Size(115,315)
  $FormSettingsCancelButton.Size=New-Object System.Drawing.Size(80,25)
  $FormSettingsCancelButton.Text="Cancel"
  $FormSettingsCancelButton.BackColor="#FFCCCCCC"
  $FormSettingsCancelButton.Add_Click( { 
    Restore-CurrentSettings
    $FormSettings.Close()
  } )
  $FormSettings.Controls.Add($FormSettingsCancelButton)

  # Form layout correction
  $InitialFormWindowState=New-Object System.Windows.Forms.FormWindowState
  $InitialFormWindowState=$Form.WindowState
  $Form.Add_Load( { $Form.WindowState=$InitialFormWindowState } )

  $Form.ResumeLayout()
  $MenuToolStrip.ResumeLayout()

  #Define the Timer for the grid refreshes
  $Timer=New-Object System.windows.Forms.Timer
  $Timer.Enabled=$False
  $TimerScheduler=New-Object System.windows.Forms.Timer
  $TimerScheduler.Enabled=$False

}

function Load-Logo {

  # Load the Hydra Splash Screen/About picture

  $FormSplashScreen=New-Object System.Windows.Forms.Form
  $FormSplashScreen.StartPosition="CenterScreen"
  $FormSplashScreen.FormBorderStyle="None"
  $FormSplashScreen.MinimizeBox=$False
  $FormSplashScreen.MaximizeBox=$False
  $FormSplashScreen.Topmost=$True
  $FormSplashScreen.Width=510
  $FormSplashScreen.Height=360
  $FormSplashScreen.Add_Shown({ Start-Sleep 3; $FormSplashScreen.Close() })

  $FormAboutPictureBox=New-Object Windows.Forms.PictureBox
  $FormAboutPictureBox.Width=510
  $FormAboutPictureBox.Height=360
  $FormAboutPictureBox.Image=$LogoImageFile
  $FormSplashScreen.Controls.Add($FormAboutPictureBox)

}
