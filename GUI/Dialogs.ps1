function BuildDialog {  # Help function for fBrowse-Folder_Modern 
  
  $SourceCode = @"
using System;
using System.Windows.Forms;
using System.Reflection;
namespace FolderSelect
{
	public class FolderSelectDialog
	{
		System.Windows.Forms.OpenFileDialog ofd = null;
		public FolderSelectDialog()
		{
			ofd = new System.Windows.Forms.OpenFileDialog();
			ofd.Filter = "Folders|\n";
			ofd.AddExtension = false;
			ofd.CheckFileExists = false;
			ofd.DereferenceLinks = true;
			ofd.Multiselect = false;
		}
		public string InitialDirectory
		{
			get { return ofd.InitialDirectory; }
			set { ofd.InitialDirectory = value == null || value.Length == 0 ? Environment.CurrentDirectory : value; }
		}
		public string Title
		{
			get { return ofd.Title; }
			set { ofd.Title = value == null ? "Select a folder" : value; }
		}
		public string FileName
		{
			get { return ofd.FileName; }
		}
		public bool ShowDialog()
		{
			return ShowDialog(IntPtr.Zero);
		}
		public bool ShowDialog(IntPtr hWndOwner)
		{
			bool flag = false;

			if (Environment.OSVersion.Version.Major >= 6)
			{
				var r = new Reflector("System.Windows.Forms");
				uint num = 0;
				Type typeIFileDialog = r.GetType("FileDialogNative.IFileDialog");
				object dialog = r.Call(ofd, "CreateVistaDialog");
				r.Call(ofd, "OnBeforeVistaDialog", dialog);
				uint options = (uint)r.CallAs(typeof(System.Windows.Forms.FileDialog), ofd, "GetOptions");
				options |= (uint)r.GetEnum("FileDialogNative.FOS", "FOS_PICKFOLDERS");
				r.CallAs(typeIFileDialog, dialog, "SetOptions", options);
				object pfde = r.New("FileDialog.VistaDialogEvents", ofd);
				object[] parameters = new object[] { pfde, num };
				r.CallAs2(typeIFileDialog, dialog, "Advise", parameters);
				num = (uint)parameters[1];
				try
				{
					int num2 = (int)r.CallAs(typeIFileDialog, dialog, "Show", hWndOwner);
					flag = 0 == num2;
				}
				finally
				{
					r.CallAs(typeIFileDialog, dialog, "Unadvise", num);
					GC.KeepAlive(pfde);
				}
			}
			else
			{
				var fbd = new FolderBrowserDialog();
				fbd.Description = this.Title;
				fbd.SelectedPath = this.InitialDirectory;
				fbd.ShowNewFolderButton = false;
				if (fbd.ShowDialog(new WindowWrapper(hWndOwner)) != DialogResult.OK) return false;
				ofd.FileName = fbd.SelectedPath;
				flag = true;
			}
			return flag;
		}
	}
	public class WindowWrapper : System.Windows.Forms.IWin32Window
	{
		public WindowWrapper(IntPtr handle)
		{
			_hwnd = handle;
		}
		public IntPtr Handle
		{
			get { return _hwnd; }
		}

		private IntPtr _hwnd;
	}
	public class Reflector
	{
		string m_ns;
		Assembly m_asmb;
		public Reflector(string ns)
			: this(ns, ns)
		{ }
		public Reflector(string an, string ns)
		{
			m_ns = ns;
			m_asmb = null;
			foreach (AssemblyName aN in Assembly.GetExecutingAssembly().GetReferencedAssemblies())
			{
				if (aN.FullName.StartsWith(an))
				{
					m_asmb = Assembly.Load(aN);
					break;
				}
			}
		}
		public Type GetType(string typeName)
		{
			Type type = null;
			string[] names = typeName.Split('.');

			if (names.Length > 0)
				type = m_asmb.GetType(m_ns + "." + names[0]);

			for (int i = 1; i < names.Length; ++i) {
				type = type.GetNestedType(names[i], BindingFlags.NonPublic);
			}
			return type;
		}
		public object New(string name, params object[] parameters)
		{
			Type type = GetType(name);
			ConstructorInfo[] ctorInfos = type.GetConstructors();
			foreach (ConstructorInfo ci in ctorInfos) {
				try {
					return ci.Invoke(parameters);
				} catch { }
			}

			return null;
		}
		public object Call(object obj, string func, params object[] parameters)
		{
			return Call2(obj, func, parameters);
		}
		public object Call2(object obj, string func, object[] parameters)
		{
			return CallAs2(obj.GetType(), obj, func, parameters);
		}
		public object CallAs(Type type, object obj, string func, params object[] parameters)
		{
			return CallAs2(type, obj, func, parameters);
		}
		public object CallAs2(Type type, object obj, string func, object[] parameters) {
			MethodInfo methInfo = type.GetMethod(func, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
			return methInfo.Invoke(obj, parameters);
		}
		public object Get(object obj, string prop)
		{
			return GetAs(obj.GetType(), obj, prop);
		}
		public object GetAs(Type type, object obj, string prop) {
			PropertyInfo propInfo = type.GetProperty(prop, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
			return propInfo.GetValue(obj, null);
		}
		public object GetEnum(string typeName, string name) {
			Type type = GetType(typeName);
			FieldInfo fieldInfo = type.GetField(name);
			return fieldInfo.GetValue(null);
		}
	}
}
"@

  $Assemblies=('System.Windows.Forms', 'System.Reflection')
  Add-Type -TypeDefinition $SourceCode -ReferencedAssemblies $Assemblies -ErrorAction Stop

}


function Read-InputBoxDialog([string]$WindowTitle="inputbox", [string]$Message="Enter the value:", [string]$DefaultText="") { 

  # Variable Type 1: Display input box and return the value entered by the user

  Add-Type -AssemblyName Microsoft.VisualBasic
  
  $form=New-Object System.Windows.Forms.Form 
  $form.Text=$WindowTitle
  $form.Size=New-Object System.Drawing.Size(410,175)
  $form.StartPosition="CenterScreen"
  $form.FormBorderStyle='FixedSingle'
  $form.ControlBox=$False
  $form.Topmost=$True
  $form.ShowInTaskbar=$True
  $form.KeyPreview=$True
  $form.Add_KeyDown( { 
    if ($_.KeyCode -eq "Enter") { 
      $Form.Tag=$TextBox.Text
      $form.Close() 
    } 
  } )
  $form.Add_KeyDown( { 
    if ($_.KeyCode -eq "Escape") { 
      $form.Tag=""
      $form.Close()
    } 
  } )

  $IconPictureBox=New-Object Windows.Forms.PictureBox
  $IconPictureBox.Width=64
  $IconPictureBox.Height=64
  $IconPictureBox.Top=30
  $IconPictureBox.Left=10
  $IconPictureBox.Image=$Icon08
  $form.Controls.Add($IconPictureBox)

  $OKButton=New-Object System.Windows.Forms.Button
  $OKButton.Location=New-Object System.Drawing.Size(220,100)
  $OKButton.Size=New-Object System.Drawing.Size(77,25)
  $OKButton.Text=" OK"
  $OKButton.Image=$Icon135
  $OKButton.TextImageRelation="ImageBeforeText"
  $OKButton.UseVisualStyleBackColor=$True
  $OKButton.Add_Click( { 
    $form.Tag=$TextBox.Text
    $form.Close() 
  } )
  $form.Controls.Add($OKButton)

  $CancelButton=New-Object System.Windows.Forms.Button
  $CancelButton.Location=New-Object System.Drawing.Size(310,100)
  $CancelButton.Size=New-Object System.Drawing.Size(77,25)
  $CancelButton.Text=" Cancel"
  $CancelButton.Image=$Icon113
  $CancelButton.TextImageRelation="ImageBeforeText"
  $CancelButton.UseVisualStyleBackColor=$True
  $CancelButton.Add_Click( { 
    $form.Tag=""
    $form.Close() 
  } )
  $form.Controls.Add($CancelButton)

  $Label=New-Object System.Windows.Forms.Label
  $Label.Location=New-Object System.Drawing.Size(80,20) 
  $Label.Size=New-Object System.Drawing.Size(300,20)
  $Label.Text=$Message
  $Label.AutoSize=$False
  $Label.AutoEllipsis=$True
  $form.Controls.Add($Label) 

  $TextBox=New-Object System.Windows.Forms.TextBox
  $TextBox.Location=New-Object System.Drawing.Size(100,55) 
  $TextBox.Size=New-Object System.Drawing.Size(283,20) 
  $TextBox.Text=$DefaultText
  $form.Controls.Add($TextBox) 
  $form.Topmost=$True
  $form.Add_Shown({ 
    $form.Activate()
    $TextBox.Focus() 
  })

  [void] $form.ShowDialog()

  return $form.Tag 

}


function Read-MultiLineInputBoxDialog([string]$WindowTitle="multilineinputbox", [string]$Message="Enter the values:", [string]$DefaultText="") {

  # Variable Type 2: Display a multi lines input box and return the value entered by the user

  Add-Type -AssemblyName System.Drawing 
  Add-Type -AssemblyName System.Windows.Forms 
      
  $label=New-Object System.Windows.Forms.Label 
  $label.Location=New-Object System.Drawing.Size(10,15)  
  $label.Size=New-Object System.Drawing.Size(480,20) 
  $label.AutoSize=$False
  $label.AutoEllipsis=$True
  $label.Text=$Message
          
  $textBox=New-Object System.Windows.Forms.TextBox  
  $textBox.Location=New-Object System.Drawing.Size(15,40)  
  $textBox.Size=New-Object System.Drawing.Size(475,200) 
  $textBox.AcceptsReturn=$True
  $textBox.AcceptsTab=$False
  $textBox.Multiline=$True
  $textBox.ScrollBars='Both'
  $textBox.Text=$DefaultText

  $IconPictureBox=New-Object Windows.Forms.PictureBox
  $IconPictureBox.Width=64
  $IconPictureBox.Height=64
  $IconPictureBox.Top=140
  $IconPictureBox.Left=510
  $IconPictureBox.Image=$Icon09
    
  $OKButton=New-Object System.Windows.Forms.Button 
  $OKButton.Location=New-Object System.Drawing.Size(510,40) 
  $OKButton.Size=New-Object System.Drawing.Size(77,25) 
  $OKButton.Text=" OK"
  $OKButton.Image=$Icon135
  $OKButton.TextImageRelation="ImageBeforeText"
  $OKButton.UseVisualStyleBackColor=$True
  $OKButton.Add_Click( { $form.Tag=$textBox.Text; $form.Close() } ) 
    
  $CancelButton=New-Object System.Windows.Forms.Button 
  $CancelButton.Location=New-Object System.Drawing.Size(510,80) 
  $CancelButton.Size=New-Object System.Drawing.Size(77,25) 
  $CancelButton.Text=" Cancel"
  $CancelButton.Image=$Icon113
  $CancelButton.TextImageRelation="ImageBeforeText"
  $CancelButton.UseVisualStyleBackColor=$True
  $CancelButton.Add_Click( { $form.Tag=""; $form.Close() } ) 
    
  $form=New-Object System.Windows.Forms.Form  
  $form.Text=$WindowTitle
  $form.Size=New-Object System.Drawing.Size(610,285) 
  $form.StartPosition="CenterScreen"
  $form.FormBorderStyle='FixedSingle'
  $form.ControlBox=$False
  $form.Topmost=$True
  $form.AcceptButton=$okButton
  $form.CancelButton=$cancelButton
  $form.ShowInTaskbar=$True
       
  $form.Controls.Add($label)
  $form.Controls.Add($textBox)
  $form.Controls.Add($OKButton)
  $form.Controls.Add($CancelButton)
  $form.Controls.Add($IconPictureBox)

  $form.Add_Shown( { $form.Activate() } ) 
  $form.ShowDialog() | Out-Null
           
  return $form.Tag 

}


function Read-OpenFileDialog([string]$WindowTitle="selectfile", [string]$InitialDirectory="C:\temp", [string]$Filter="All files (*.*)|*.*", [switch]$AllowMultiSelect=$False) {  

  # Variable Type 3: Display an Open File Dialog window and return the file selected by the user 

  Add-Type -AssemblyName System.Windows.Forms 
  $openFileDialog=New-Object System.Windows.Forms.OpenFileDialog 
  $openFileDialog.Title=$WindowTitle
  if (![string]::IsNullOrWhiteSpace($InitialDirectory)) { 
	$openFileDialog.InitialDirectory=$InitialDirectory 
  } 
  $openFileDialog.Filter=$Filter
  if ($AllowMultiSelect) { 
	$openFileDialog.MultiSelect=$True 
  } 
  $openFileDialog.ShowHelp=$True    # Without this line the ShowDialog() function may hang depending on system configuration and running from console vs. ISE.     
  $openFileDialog.ShowDialog() > $null
  if ($AllowMultiSelect) { 
	return $openFileDialog.Filenames 
  } 
  else { 
	return $openFileDialog.Filename 
  }

}


function fBrowse-Folder_Modern([string]$SelectedPath="C:\temp") {

  # Variable Type 4: Display an Select Folder Dialog window and return the folder selected by the user

  BuildDialog
  $FolderBrowser=New-Object FolderSelect.FolderSelectDialog
  if ($SelectedPath -ne $Null) {
	$FolderBrowser.InitialDirectory = $SelectedPath
  }
  [void]$FolderBrowser.ShowDialog()
  return $FolderBrowser.FileName

}


function Read-ComboBoxDialog([string]$WindowTitle="combobox", [string]$Message="Choose the value:", $Entries) {

  # Variable Type 5: Display a ComboBox dialog box and return the value selected by the user

  function fForm-OnLoad {
    $Options=$Null
    if (@($Entries).Count -eq 1) {  # One entry only: could be a file
      $SeqPath=$(Split-Path $SequenceFullPath -Parent)
      if (Test-Path $Entries -ErrorAction SilentlyContinue) {
        $Options=Get-Content $Entries
	  }
      elseif (Test-Path $(Join-Path -Path $SeqPath -ChildPath $Entries) -ErrorAction SilentlyContinue) {
        $Options=Get-Content $(Join-Path -Path $SeqPath -ChildPath $Entries)
      }
      elseif (Test-Path $(Join-Path -Path $HydraBinPath -ChildPath $Entries) -ErrorAction SilentlyContinue) {
        $Options=Get-Content $(Join-Path -Path $HydraBinPath -ChildPath $Entries)
      }
    }
    if ($Options -eq $Null) {
      $Options=$Entries -split ","
	}
    foreach($Entry in $Options) { 
      if (!([string]::IsNullOrWhiteSpace($Entry))) { [void]$VarComboBox.Items.Add($Entry.Trim()) }
    }
    $VarComboBox.SelectedIndex=0
  }

  $Form=New-Object System.Windows.Forms.Form  
  $Form.Text=$WindowTitle
  $Form.Size=New-Object System.Drawing.Size(450,160) 
  $Form.FormBorderStyle='FixedSingle'
  $Form.StartPosition="CenterScreen"
  $Form.ControlBox=$False
  $Form.Topmost=$True
  $Form.ShowInTaskbar=$True

  $IconPictureBox=New-Object Windows.Forms.PictureBox
  $IconPictureBox.Width=64
  $IconPictureBox.Height=64
  $IconPictureBox.Top=25
  $IconPictureBox.Left=10
  $IconPictureBox.Image=$Icon12
  $Form.Controls.Add($IconPictureBox)

  $Label=New-Object System.Windows.Forms.Label 
  $Label.Location=New-Object System.Drawing.Size(80,15)  
  $Label.Size=New-Object System.Drawing.Size(330,20) 
  $Label.AutoSize=$False
  $Label.AutoEllipsis=$True
  $Label.Text=$Message
  $Form.Controls.Add($Label)
  
  $VarComboBox=New-Object System.Windows.Forms.ComboBox
  $VarComboBox.Location=New-Object System.Drawing.Point(90,45)
  $VarComboBox.Size=New-Object System.Drawing.Size(330, 310)
  $VarComboBox.DropDownStyle='DropDownList'
  $VarComboBox.AutoCompleteMode='None'
  $VarComboBox.AutoCompleteSource='ListItems'
  $Form.Controls.Add($VarComboBox)
   
  $OKButton=New-Object System.Windows.Forms.Button 
  $OKButton.Location=New-Object System.Drawing.Size(150,90) 
  $OKButton.Size=New-Object System.Drawing.Size(77,25) 
  $OKButton.Text=" OK"
  $OKButton.Image=$Icon135
  $OKButton.TextImageRelation="ImageBeforeText"
  $OKButton.UseVisualStyleBackColor=$True
  $OKButton.Add_Click( { $Form.Tag=$VarComboBox.Text; $Form.Close() } ) 
  $Form.Controls.Add($OKButton)
      
  $CancelButton=New-Object System.Windows.Forms.Button 
  $CancelButton.Location=New-Object System.Drawing.Size(265,90) 
  $CancelButton.Size=New-Object System.Drawing.Size(77,25) 
  $CancelButton.Text=" Cancel"
  $CancelButton.Image=$Icon113
  $CancelButton.TextImageRelation="ImageBeforeText"
  $CancelButton.UseVisualStyleBackColor=$True
  $CancelButton.Add_Click( { $Form.Tag = ""; $Form.Close() } ) 
  $Form.Controls.Add($CancelButton)

  $Form.Add_Load( { fForm-OnLoad } )

  [void]$Form.ShowDialog()

  return $Form.Tag

}


function Read-MultiCheckboxList([string]$WindowTitle="multicheckbox", [string]$Message="Select the values", $Entries) {

  # Variable Type 6: Display a list of Checkboxes and return the value(s) selected by the user
	
  function Form-OnLoad {
    $Options=$Null
    if (@($Entries).Count -eq 1) {  # One entry only: could be a file
      $SeqPath=$(Split-Path $SequenceFullPath -Parent)
      if (Test-Path $Entries -ErrorAction SilentlyContinue) {
        $Options=Get-Content $Entries
      }
      elseif (Test-Path $(Join-Path -Path $SeqPath -ChildPath $Entries) -ErrorAction SilentlyContinue) {
        $Options=Get-Content $(Join-Path -Path $SeqPath -ChildPath $Entries)
      }
      elseif (Test-Path $(Join-Path -Path $HydraBinPath -ChildPath $Entries) -ErrorAction SilentlyContinue) {
        $Options=Get-Content $(Join-Path -Path $HydraBinPath -ChildPath $Entries)
      }
    }
    if ($Options -eq $Null) {
      $Options=$Entries -split ","
	}
    $lcOptions.BeginUpdate()
    foreach ($Entry in $Options) {
      if (!([string]::IsNullOrWhiteSpace($Entry))) {
        $lcOptions.Items.Add($Entry.Trim()) | Out-Null
      }
    }
	$lcOptions.EndUpdate()
    $Y=($lcOptions.Items.Count)*18
    if ($Y -gt 350) { $Y=350 }
    if ($Y -lt 170) { $Y=170 }
    $lcOptions.Size=New-Object System.Drawing.Size 300, $Y
    $form.Size=New-Object System.Drawing.Size 440, $($Y+110)
    $okButton.Location=New-Object System.Drawing.Size(15, $($form.Height - 70))
    $cancelButton.Location=New-Object System.Drawing.Size(110, $($form.Height - 70))
  }
	
  function Select-All {
    if ($rAll.Checked -eq $true) {
      for ($i=0; $i -lt $lcOptions.Items.Count; $i++) {
	    $lcOptions.SetItemchecked($i,$True)
	  }
	}
  }

  function UnSelect-All {
    if ($rNone.Checked -eq $true) {
	  for ($i=0; $i -lt $lcOptions.Items.count; $i++) {
        $lcOptions.SetItemchecked($i,$False)
      }
	}
  }
	
  $label=New-Object System.Windows.Forms.Label 
  $label.Location=New-Object System.Drawing.Size(10,10)  
  $label.Size=New-Object System.Drawing.Size(280,20) 
  $label.AutoSize=$False
  $label.AutoEllipsis=$True
  $label.Text=$Message

  $IconPictureBox=New-Object Windows.Forms.PictureBox
  $IconPictureBox.Width=64
  $IconPictureBox.Height=64
  $IconPictureBox.Top=110
  $IconPictureBox.Left=335
  $IconPictureBox.Image=$Icon11

  $lcOptions=New-Object System.Windows.Forms.CheckedListBox
  $lcOptions.Font=New-Object System.Drawing.Font("Arial", 9, 0, 3, 1)
  $lcOptions.FormattingEnabled=$True
  $lcOptions.Location=New-Object System.Drawing.Point(12,30)
  $lcOptions.Size=New-Object System.Drawing.Size(300,350)
  $lcOptions.TabIndex=0
  $lcOptions.CheckOnClick=$True
  
  $OKButton=New-Object System.Windows.Forms.Button 
  $OKButton.Size=New-Object System.Drawing.Size(77,25) 
  $OKButton.Text=" OK"
  $OKButton.Image=$Icon135
  $OKButton.TextImageRelation="ImageBeforeText"
  $OKButton.UseVisualStyleBackColor=$True
  $OKButton.Add_Click({ 
    $form.Tag=@()
	foreach ($item in $lcOptions.CheckedItems) {
      $form.Tag+=$item.ToString()
    }
    if ($form.Tag.Count -eq 0) {
      [string]$form.Tag="NONE"
    }
	$form.Close() 
  }) 

  $CancelButton=New-Object System.Windows.Forms.Button 
  $CancelButton.Size=New-Object System.Drawing.Size(77,25) 
  $CancelButton.Text=" Cancel"
  $CancelButton.Image=$Icon113
  $CancelButton.TextImageRelation="ImageBeforeText"
  $CancelButton.UseVisualStyleBackColor=$True
  $CancelButton.Add_Click( { $form.Tag=""; $form.Close() } ) 
	
  $rNone=New-Object System.Windows.Forms.RadioButton
  $rNone.Location=New-Object System.Drawing.Point(330,60)
  $rNone.Size=New-Object System.Drawing.Size(100,25)
  $rNone.Text="Unselect All"
  $rNone.Checked=$True
  $rNone.UseVisualStyleBackColor=$True
  $rNone.add_Click( { UnSelect-All } )

  $rAll=New-Object System.Windows.Forms.RadioButton
  $rAll.Location=New-Object System.Drawing.Point(330,30)
  $rAll.Size=New-Object System.Drawing.Size(100,25)
  $rAll.TabStop=$True
  $rAll.Text="Select All"
  $rAll.UseVisualStyleBackColor=$True
  $rAll.add_Click( { Select-All } )

  $form=New-Object System.Windows.Forms.Form  
  $form.Text=$WindowTitle
  $form.Size=New-Object System.Drawing.Size(440,420) 
  $form.FormBorderStyle='FixedSingle'
  $form.StartPosition="CenterScreen"
  $form.AutoSizeMode='GrowAndShrink'
  $form.Topmost=$True
  $form.ControlBox=$False
  $form.AcceptButton=$okButton
  $form.CancelButton=$cancelButton
  $form.ShowInTaskbar=$True
   
  $form.Controls.Add($label) 
  $form.Controls.Add($lcOptions)
  $form.Controls.Add($OKButton) 
  $form.Controls.Add($CancelButton) 
  $form.Controls.Add($rNone)
  $form.Controls.Add($rAll)
  $form.Controls.Add($IconPictureBox)
  
  $form.Add_Load( { Form-OnLoad } )

  $form.Add_Shown( { $form.Activate() } ) 
  $form.ShowDialog() | out-null         
  
  return $form.Tag

}


function Read-DateTimePicker([string]$WindowTitle="scheduler") {

  # Variable Type 7: Display a Date/Time picker dialog and return the values selected by the user

  function fSetDate-Time {
	$Date=$DatePicker.Value.ToString("MM/dd/yyyy") 
	$Time=$TimePicker.Value.ToString("HH:mm:ss") 
	$form1.Tag=[datetime]"$Date $Time"
	$CurrentDate=Get-Date
	if ($form1.Tag -le $CurrentDate) {
	  [System.Windows.Forms.MessageBox]::Show("Date/Time cannot be older than current date/time" , "Error")
	  return
	}
	$form1.Close()
  }

  $form1=New-Object System.Windows.Forms.Form
  $form1.ClientSize="350,170"
  $form1.Text=$(if ($WindowTitle -eq "") { "Scheduler" } else { $WindowTitle } )
  $form1.StartPosition="CenterScreen"
  $form1.FormBorderStyle="FixedDialog"
  $form1.ControlBox=$False
  $form1.Topmost=$True
  $form1.ShowInTaskbar=$True
  $form1.KeyPreview=$True
  $form1.Add_Shown( { $form1.Activate() } ) 
  $form1.Add_KeyDown( { if ($_.KeyCode -eq "Escape") {$form1.Close() } } )

  $Label=New-Object System.Windows.Forms.Label 
  $Label.Location=New-Object System.Drawing.Size(70,15)  
  $Label.Size=New-Object System.Drawing.Size(330,20) 
  $Label.AutoSize=$False
  $Label.AutoEllipsis=$True
  $Label.Text="Pick a Date and Time for the scheduler:"
  $form1.Controls.Add($Label)

  $IconPictureBox=New-Object Windows.Forms.PictureBox
  $IconPictureBox.Width=64
  $IconPictureBox.Height=64
  $IconPictureBox.Top=40
  $IconPictureBox.Left=5
  $IconPictureBox.Image=$Icon13
  $form1.Controls.Add($IconPictureBox)
  
  $ldate=New-Object System.Windows.Forms.Label	
  $ldate.Location="80,50"
  $ldate.Size="40,23"
  $ldate.Text="Date:"
  $form1.Controls.Add($ldate)	
	
  $lTime=New-Object System.Windows.Forms.Label
  $lTime.Location="80,90"
  $lTime.Size="40,23"
  $lTime.Text="Time:"
  $form1.Controls.Add($lTime)	

  $OKButton=New-Object System.Windows.Forms.Button
  $OKButton.Location="150,130"
  $OKButton.Size="77,25"
  $OKButton.Text=" OK"
  $OKButton.Image=$Icon135
  $OKButton.TextImageRelation="ImageBeforeText"
  $OKButton.UseVisualStyleBackColor=$True
  $OKButton.UseVisualStyleBackColor=$True
  $OKButton.add_Click( { fSetDate-Time } )
  $form1.Controls.Add($OKButton)

  $CancelButton=New-Object System.Windows.Forms.Button
  $CancelButton.Location="250,130"
  $CancelButton.Size="77,25"
  $CancelButton.Text=" Cancel"
  $CancelButton.Image=$Icon113
  $CancelButton.TextImageRelation="ImageBeforeText"
  $CancelButton.UseVisualStyleBackColor=$True
  $CancelButton.UseVisualStyleBackColor=$True
  $CancelButton.add_Click( { $Form1.Tag="" ; $form1.Close() } )
  $form1.Controls.Add($CancelButton)

  $DatePicker=New-Object System.Windows.Forms.DateTimePicker
  $datePicker.Font=New-Object System.Drawing.Font("Microsoft Sans Serif", 9, 0, 3, 0)
  $datePicker.Location="120,46"
  $datePicker.Size="210,22"
  $form1.Controls.Add($datePicker)

  $TimePicker=New-Object System.Windows.Forms.DateTimePicker
  $TimePicker.Font=New-Object System.Drawing.Font("Microsoft Sans Serif", 9, 0, 3, 0)
  $TimePicker.Format=4
  $TimePicker.Location="120,86"
  $TimePicker.ShowUpDown=$True
  $TimePicker.Size="85,21"
  $form1.Controls.Add($TimePicker)

  $form1.AcceptButton=$OKButton
  $form1.CancelButton=$CancelButton
  $form1.ShowDialog()| Out-Null
  
  return $form1.Tag
  
} 


function Read-Credentials($Message="Enter the credentials:") {

  # Variable Type 8: Display a Credentials dialog box and return the values entered by the user
  
  try {
    $Cred=Get-Credential -Message $Message
  }

  catch {
    return ""
  }

  return $Cred

}


function Read-InputDialogBoxSecret($Header="secretinputbox", $Message="Enter the secret information below:") {

  # Variable Type 9: Display a Secret input box (all characters are replaced by x) and return the value entered by the user

  $form=New-Object System.Windows.Forms.Form 
  $form.Text=$Header
  $form.Size=New-Object System.Drawing.Size(310,175) 
  $form.StartPosition="CenterScreen"
  $form.FormBorderStyle='FixedSingle'
  $form.ControlBox=$False
  $form.Topmost=$True
  $form.ShowInTaskbar=$True
  $form.KeyPreview=$True
  $form.Add_KeyDown( { 
    if ($_.KeyCode -eq "Enter") { 
      $Form.Tag=$MaskedTextBox.Text
      $form.Close() 
    } 
  } )
  $form.Add_KeyDown( { 
    if ($_.KeyCode -eq "Escape") { 
      $form.Tag=""
      $form.Close()
    } 
  } )

  $IconPictureBox=New-Object Windows.Forms.PictureBox
  $IconPictureBox.Width=64
  $IconPictureBox.Height=64
  $IconPictureBox.Top=30
  $IconPictureBox.Left=10
  $IconPictureBox.Image=$Icon10
  $form.Controls.Add($IconPictureBox)

  $OKButton=New-Object System.Windows.Forms.Button
  $OKButton.Location=New-Object System.Drawing.Size(120,100)
  $OKButton.Size=New-Object System.Drawing.Size(77,25)
  $OKButton.Text=" OK"
  $OKButton.Image=$Icon135
  $OKButton.TextImageRelation="ImageBeforeText"
  $OKButton.UseVisualStyleBackColor=$True
  $OKButton.Add_Click( { 
    $form.Tag=$MaskedTextBox.Text
    $form.Close() 
  } )
  $form.Controls.Add($OKButton)

  $CancelButton=New-Object System.Windows.Forms.Button
  $CancelButton.Location=New-Object System.Drawing.Size(210,100)
  $CancelButton.Size=New-Object System.Drawing.Size(77,25)
  $CancelButton.Text=" Cancel"
  $CancelButton.Image=$Icon113
  $CancelButton.TextImageRelation="ImageBeforeText"
  $CancelButton.UseVisualStyleBackColor=$True
  $CancelButton.Add_Click( { 
    $form.Tag=""
    $form.Close() 
  } )
  $form.Controls.Add($CancelButton)

  $Label=New-Object System.Windows.Forms.Label
  $Label.Location=New-Object System.Drawing.Size(80,20) 
  $Label.Size=New-Object System.Drawing.Size(200,20) 
  $Label.Text=$Message
  $Label.AutoSize=$False
  $Label.AutoEllipsis=$True
  $form.Controls.Add($Label) 

  $MaskedTextBox=New-Object System.Windows.Forms.MaskedTextBox
  $MaskedTextBox.PasswordChar='x'
  $MaskedTextBox.Location=New-Object System.Drawing.Size(100,55) 
  $MaskedTextBox.Size=New-Object System.Drawing.Size(183,20) 
  $form.Controls.Add($MaskedTextBox) 
  $form.Topmost=$True
  $form.Add_Shown({ 
    $form.Activate()
    $MaskedTextBox.Focus() 
  })
  [void] $form.ShowDialog()

  return $form.Tag

}


function Read-SecurityCode($SecurityCode, $SequenceName) {

  if ($SecurityCode.Length -gt 20) {
    $SecurityCode=$SecurityCode.Substring(0,20)
  }
  $Script:Tries=1

  $ButtonOK_Click={
      if ($TextBox.Text -ceq $SecurityCode) {
      $SecForm.Close()
      $SecForm.Tag="OK"
    }
    elseif ($Tries -ge 3) {
      $SecForm.Close()
      [void][System.Windows.Forms.MessageBox]::Show("Too many wrong tries: the sequence start has been cancelled.", "Sequence", [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
      $SecForm.Tag="KO"
    }
    else {
      $Script:Tries++
      [void][System.Windows.Forms.MessageBox]::Show("The Code does not match: be sure to respect the case.", "Sequence", [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
  }

  $TextBox_KeyPressHandler=[System.Windows.Forms.KeyPressEventHandler]{
    $Key=$_
    if ($Key.KeyChar -eq 13) { Invoke-Command $ButtonOK_Click }
  }

  $SecForm=New-Object System.Windows.Forms.Form
  $SecForm.Size=New-Object System.Drawing.Size(360,240)   
  $SecForm.Font=New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,0,3,0)
  $SecForm.Text="$SequenceName"
  $SecForm.StartPosition="CenterParent"
  $SecForm.FormBorderStyle='FixedSingle'
  $SecForm.ControlBox=$False
  $SecForm.Topmost=$True
  $SecForm.ShowInTaskbar=$True

  $IconPictureBox=New-Object Windows.Forms.PictureBox
  $IconPictureBox.Width=64
  $IconPictureBox.Height=64
  $IconPictureBox.Top=50
  $IconPictureBox.Left=15
  $IconPictureBox.Image=$Icon07
  $SecForm.Controls.Add($IconPictureBox)

  $Label=New-Object System.Windows.Forms.Label
  $Label.Location=New-Object System.Drawing.Size(80, 26)
  $Label.Size=New-Object System.Drawing.Size(258, 20)
  $Label.Text="Please enter the code below to start the Sequence"
  $SecForm.Controls.Add($Label)
  
  $LabelSeq=New-Object System.Windows.Forms.Label
  $LabelSeq.Location=New-Object System.Drawing.Size(80, 46)
  $LabelSeq.Size=New-Object System.Drawing.Size(248, 20)
  $LabelSeq.AutoSize=$False
  $LabelSeq.AutoEllipsis=$True
  $LabelSeq.Text="""$SequenceName"":"
  $SecForm.Controls.Add($LabelSeq)

  $Panel=New-Object System.Windows.Forms.Panel
  $Panel.Location=New-Object System.Drawing.Size(110, 82)
  $Panel.Size=New-Object System.Drawing.Size(176, 32)
  $Panel.BackColor=[System.Drawing.Color]::FromArgb(255,153,180,209)
  $Panel.BorderStyle=1
  $SecForm.Controls.Add($Panel)

  $LabelCode=New-Object System.Windows.Forms.Label
  $LabelCode.Font=New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,1,3,0)
  $LabelCode.Location=New-Object System.Drawing.Size(2, 2)
  $LabelCode.Size=New-Object System.Drawing.Size(170, 23)
  $LabelCode.Text=$SecurityCode
  $LabelCode.TextAlign=32
  $Panel.Controls.Add($LabelCode)

  $TextBox=New-Object System.Windows.Forms.TextBox
  $TextBox.Location=New-Object System.Drawing.Size(110, 126)
  $TextBox.Size=New-Object System.Drawing.Size(176, 20)
  $TextBox.Text=""
  $TextBox.TextAlign=2
  $TextBox.TabIndex=1
  $SecForm.Controls.Add($TextBox)
  $TextBox.Add_KeyPress( $TextBox_KeyPressHandler )

  $ButtonStart=New-Object System.Windows.Forms.Button
  $ButtonStart.Location=New-Object System.Drawing.Size(120, 165)
  $ButtonStart.Size=New-Object System.Drawing.Size(77, 25)
  $ButtonStart.Text=" Start"
  $ButtonStart.UseVisualStyleBackColor=$True
  $ButtonStart.Add_Click( $ButtonOK_Click )
  $ButtonStart.Image=$Icon135
  $ButtonStart.TextImageRelation="ImageBeforeText"
  $SecForm.Controls.Add($ButtonStart)

  $ButtonCancel=New-Object System.Windows.Forms.Button
  $ButtonCancel.Location=New-Object System.Drawing.Size(220, 165)
  $ButtonCancel.Size=New-Object System.Drawing.Size(77, 25)
  $ButtonCancel.Text=" Cancel"
  $ButtonCancel.Image=$Icon113
  $ButtonCancel.TextImageRelation="ImageBeforeText"
  $ButtonCancel.UseVisualStyleBackColor=$True
  $ButtonCancel.Add_Click( {
    $SecForm.Hide()
    $SecForm.BringToFront()
    $SecForm.Close()
    [void][System.Windows.Forms.MessageBox]::Show("The sequence start has been cancelled.", "Sequence", [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
    $SecForm.Tag="KO"
  } )
  $SecForm.Controls.Add($ButtonCancel)
  
  $SecForm.ShowDialog() | Out-Null

  return $SecForm.Tag

}


function Read-StartSequence($SequenceName) {

  $SeqStartForm=New-Object System.Windows.Forms.Form
  $SeqStartForm.Width=320  
  $SeqStartForm.Height=170
  $SeqStartForm.Font=New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,0,3,0)
  $SeqStartForm.Text="$SequenceName"
  $SeqStartForm.StartPosition="CenterParent"
  $SeqStartForm.FormBorderStyle='FixedSingle'
  $SeqStartForm.ControlBox=$False
  $SeqStartForm.Topmost=$True
  $SeqStartForm.ShowInTaskbar=$True

  $IconPictureBox=New-Object Windows.Forms.PictureBox
  $IconPictureBox.Width=64
  $IconPictureBox.Height=64
  $IconPictureBox.Top=30
  $IconPictureBox.Left=10
  $IconPictureBox.Image=$Icon14
  $SeqStartForm.Controls.Add($IconPictureBox)

  $Label=New-Object System.Windows.Forms.Label
  $Label.Location=New-Object System.Drawing.Size(80, 26)
  $Label.Size=New-Object System.Drawing.Size(230, 20)
  $Label.Text="Do you want to proceed with the Sequence"
  $SeqStartForm.Controls.Add($Label)
  
  $LabelSeq=New-Object System.Windows.Forms.Label
  $LabelSeq.Location=New-Object System.Drawing.Size(80, 46)
  $LabelSeq.Size=New-Object System.Drawing.Size(220, 20)
  $LabelSeq.Text="""$SequenceName"":"
  $LabelSeq.AutoSize=$False
  $LabelSeq.AutoEllipsis=$True
  $SeqStartForm.Controls.Add($LabelSeq)

  $ButtonStart=New-Object System.Windows.Forms.Button
  $ButtonStart.Location=New-Object System.Drawing.Size(90, 85)
  $ButtonStart.Size=New-Object System.Drawing.Size(77, 25)
  $ButtonStart.Text=" Start"
  $ButtonStart.Image=$Icon135
  $ButtonStart.TextImageRelation="ImageBeforeText"
  $ButtonStart.UseVisualStyleBackColor=$True
  $ButtonStart.Add_Click( {
    $SeqStartForm.Close()
    $SeqStartForm.Tag="OK"
  } )
  $SeqStartForm.Controls.Add($ButtonStart)

  $ButtonCancel=New-Object System.Windows.Forms.Button
  $ButtonCancel.Location=New-Object System.Drawing.Size(200, 85)
  $ButtonCancel.Size=New-Object System.Drawing.Size(77, 25)
  $ButtonCancel.Text=" Cancel"
  $ButtonCancel.Image=$Icon113
  $ButtonCancel.TextImageRelation="ImageBeforeText"
  $ButtonCancel.UseVisualStyleBackColor=$True
  $ButtonCancel.Add_Click( { 
    $SeqStartForm.Hide()
    $SeqStartForm.BringToFront()
    $SeqStartForm.Close()
    [void][System.Windows.Forms.MessageBox]::Show("The sequence start has been cancelled.", "Sequence", [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
    $SeqStartForm.Tag="KO"
  } )
  $SeqStartForm.Controls.Add($ButtonCancel)
  
  $SeqStartForm.ShowDialog() | Out-Null

  return $SeqStartForm.Tag

}


function fSend-Mail ($SMTPServer, $From, $To, $ReplyTo, $Subject, $Body, $HTML, $Attachments) {

  # Send email function

  $SMTPObject=New-Object Net.Mail.SmtpClient($SMTPServer)
  $MailMessage=New-Object Net.Mail.MailMessage
  $MailMessage.From=$From
  $MailMessage.To.Add($To)
  if ($ReplyTo -ne $Null) { 
    $MailMessage.ReplyTo=$ReplyTo 
  }
  $MailMessage.Subject=$Subject
  $MailMessage.IsBodyHTML=$HTML
  $MailMessage.Priority="Normal"
  $MailMessage.Body=$Body
  if ($Attachments -ne $Null) {
    for ($i=0;$i -lt $Attachments.Count;$i++) {
      $Attachment=New-Object Net.Mail.Attachment($Attachments[$i])
      $MailMessage.Attachments.Add($Attachment)
    }
  }

  $SMTPObject.Send($MailMessage)

}
