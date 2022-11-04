<#
.SYNOPSIS
    An easier function-based way to create/manage windows forms in Powershell
    
.DESCRIPTION
    A Form-Building Module originally written by Dave Carroll and expanded upon by Chris Goode.

    This module allows for Windows GUI forms to be built by using a custom set of cmdlets/functions and passing parameters to the functions.
    The functions will then instanciate the desired forms, build them, draw them, and the display them with the appropriate method call from 
    the main module.  
    
.EXAMPLE
    See function-specific examples for more information
    
.NOTES
    Guaranteed to be the best free Error Handler Module or DOUBLE your money back! YOU CAN'T LOSE!
    
    I believe in giving credit where it is due.  If you are due credit for some of the functions/code contained in this module
    PLEASE email me at chrisgoode83@gmail.com and I will update the respective module and publish an update.

.LINK
    Chris Goode :   https://github.com/cgoode01/module-FormShortcuts
    Dave Carroll:   https://thedavecarroll.com/powershell/windows-forms/
    Ondrej Sebela:  https://doitpsway.com/extracting-html-table-from-a-web-page-or-html-file-and-converting-it-into-powershell-object
                    https://github.com/ztrhgf/useful_powershell_functions/blob/master/ConvertFrom-HTMLTable.ps1
    
#>


Write-Debug "......Function New-WindowsForm"
Function New-WindowsForm {
    [CmdLetBinding()]
    [OutputType([System.Windows.Forms.Form])]
    Param(
        [string]$Name,
        [int]$Width,
        [int]$Height,
        [string]$startPosition,
        $FormIcon
    )

    Try {
        $WindowsForm = [System.Windows.Forms.Form]::new()
        $WindowsForm.Name = ($Name -Replace '[^a-zA-Z]','') + 'WindowsForm'
        $WindowsForm.Text = $Name
        $WindowsForm.ClientSize = [System.Drawing.Size]::new($Width,$Height)
        $WindowsForm.DataBindings.DefaultDataSourceUpdateMode = 0
        #$WindowsForm.AutoSize = $true
        $WindowsForm.AutoSizeMode = 'GrowAndShrink'
        $WindowsForm.AutoScroll = $true
        $WindowsForm.Margin = 5
        If (-Not $FormIcon) {
            $WindowsForm.Icon = $icon
        }
        Else {
            $WindowsForm.icon = $FormIcon
        }#>
        if ($startPosition) {
            $WindowsForm.StartPosition = $startPosition
        }
        Else {
            $WindowsForm.StartPosition = 'CenterScreen'
        }#>
        $WindowsForm.WindowState = [System.Windows.Forms.FormWindowState]::new()
        $WindowsForm
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function New-TabForm"
Function New-TabForm {
    [CmdLetBinding()]
    [OutputType([System.Windows.Forms.Form])]
    Param(
        [string]$Name,
        [int]$Width,
        [int]$Height
        #[switch]$NoIcon
    )

    Try {
        $TabForm = [System.Windows.Forms.Form]::new()
        $TabForm.Name = ($Name -Replace '[^a-zA-Z]','') + 'TabForm'
        $TabForm.Text = $Name
        $TabForm.ClientSize = [System.Drawing.Size]::new($Width,$Height)
        $TabForm.DataBindings.DefaultDataSourceUpdateMode = 0
        #$WindowsForm.AutoSize = $true
        $TabForm.AutoSizeMode = 'GrowAndShrink'
        $TabForm.AutoScroll = $true
        $TabForm.Margin = 5
        $TabForm.WindowState = [System.Windows.Forms.FormWindowState]::new()
        $TabForm
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function New-TabPage"
Function New-TabPage {
    [CmdLetBinding()]
    [OutputType([System.Windows.Forms.Form])]
    Param(
        [string]$Name
        #[int]$Width,
        #[int]$Height,
        #[switch]$NoIcon
    )

    Try {
        $TabPage = [System.Windows.Forms.Form]::new()
        $TabPage.Name = ($Name -Replace '[^a-zA-Z]','') + 'TabPage'
        $TabPage.Text = $Name
        #$WindowsForm.ClientSize = [System.Drawing.Size]::new($Width,$Height)
        $TabPage.DataBindings.DefaultDataSourceUpdateMode = 0
        #$WindowsForm.AutoSize = $true
        $TabPage.AutoSizeMode = 'GrowAndShrink'
        $TabPage.AutoScroll = $true
        $TabPage.Margin = 5
        $TabPage.WindowState = [System.Windows.Forms.FormWindowState]::new()
        #if ($NoIcon) {
        #    $WindowsForm.ShowIcon = $false
        #}
        Return $TabPage
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function New-DrawingFont"
Function New-DrawingFont {
    [CmdLetBinding()]
    Param(
        [string]$Name,
        [single]$Size,
        [System.Drawing.FontStyle]$Style
    )
    Try {
        [System.Drawing.Font]::new($Name,$Size,$Style,'Point',0)
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function New-DrawingColor"
Function New-DrawingColor {
    [CmdLetBinding()]
    Param()
    Try {
        [System.Drawing.Color]::FromArgb(255,0,102,204)
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function New-FormLabel"
Function New-FormLabel {
    [CmdLetBinding()]
    [OutputType([System.Windows.Forms.Label])]
    Param(
        [string]$Name,
        [string]$Text,
        [string]$Font,
        [int]$FontSize = 10,
        [string]$FontStyle = "Regular",
        [int]$Index,
        [int]$Width,
        [int]$Height,
        [int]$DrawX,
        [int]$DrawY,
        [string]$textalign
    )
    Try {
        $FormLabel = [System.Windows.Forms.Label]::new()
        $FormLabel.Name = ($Name -Replace '[^a-zA-Z]','') + 'FormLabel'
        If ($Text) { $FormLabel.Text = $Text } Else { $FormLabel.Text = $Name }
        If ($Font) { $TextBox.Font = [System.Drawing.Font]::new("$Font",$FontSize,[System.Drawing.FontStyle]::$FontStyle) }
        $FormLabel.TabIndex = $Index
        $FormLabel.Size = [System.Drawing.Size]::new($Width,$Height)
        $FormLabel.Location = [System.Drawing.Point]::new($DrawX,$DrawY)
        $FormLabel.DataBindings.DefaultDataSourceUpdateMode = 0
        If ($textalign) {
            ## Valid Options: BottomCenter, BottomLeft, BottomRight, MiddleCenter, MiddleLeft, MiddleRight, TopCenter, TopLeft, TopRight
            $FormLabel.textalign = $textalign
        }
        $FormLabel
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function New-FormTextBox"
Function New-FormTextBox {
    [CmdLetBinding()]
    [OutputType([System.Windows.Forms.TextBox ])]
    Param(
        [string]$Name,
        [string]$Text,
        [int]$Index,
        [int]$Width,
        [int]$Height,
        [int]$DrawX,
        [int]$DrawY,
        [string]$Scrollbars = 'None',               #   None, Vertical, Horizontal, Both
        [string]$BackColor = '',
        [string]$BorderStyle = 'Fixed3D',
        [string]$Font = "Microsoft Sans Serif",
        [string]$FontColor = "Black",
        [int]$FontSize = 8,
        [string]$TextAlign = "Left",                #   Left, Right, Center
        [switch]$Bold,
        [switch]$Italic,
        [switch]$Underline,
        [switch]$Strikethru,
        [switch]$readOnly,
        [switch]$wordWrap
    )
    Try {
        $TextBox = [System.Windows.Forms.TextBox]::new()
        $TextBox.Name = ($Name -Replace '[^a-zA-Z]','') + 'FormTextBox'
        Write-Debug "Build FormTextBox Named:"
        Write-Debug $TextBox.Name
        $TextBox.Text = $Text
        $TextBox.TabIndex = $Index
        $TextBox.Size = [System.Drawing.Size]::new($Width,$Height)
        $TextBox.Location = [System.Drawing.Point]::new($DrawX,$DrawY)
        $TextBox.DataBindings.DefaultDataSourceUpdateMode = 0
        $TextBox.ScrollBars = $ScrollBars
        $TextBox.BorderStyle = $BorderStyle
        $TextBox.TextAlign = $TextAlign
        $TextBox.BackColor = $BackColor
        $TextBox.ForeColor = $FontColor
        If ($Bold)          { $FontStyle = "Bold" }
        If ($Italic)        { $FontStyle = "Italic" }
        If ($Underline)     { $FontStyle = "Underline" }
        If ($StrikeThru)    { $FontStyle = "Strikethru" }
        If ($null -eq $FontStyle) { $FontStyle = "Regular" }
        $TextBox.Font = [System.Drawing.Font]::new("$Font",$FontSize,[System.Drawing.FontStyle]::$FontStyle)
        If ($readOnly) { $TextBox.ReadOnly = $true }
        If ($wordWrap) {
            $TextBox.WordWrap = $true
            $TextBox.MultiLine = $true
        }
        $TextBox.MaxLength = 4000000
        
        $TextBox
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function New-FormRadioButton"
Function New-FormRadioButton {
    [CmdLetBinding()]
    Param(
        [string]$Name,
        [string]$Text,
        [int]$Width,
        [int]$Height,
        [int]$DrawX,
        [int]$DrawY,
        [bool]$Checked = $false,
        [switch]$Disabled
    )
    Write-Debug ""
    $timestamp = Get-Date
    Write-Debug $timestamp
    Write-Debug "Execute Function New-FormRadioButton"
    Try {  
        $RadioButton = [System.Windows.Forms.RadioButton]::new()
        $RadioButton.Name = ($Name -Replace '[^a-zA-Z]','') + 'RadioButton'
        Write-Debug "Build RadioButton Named:"
        Write-Debug $RadioButton.Name

        $RadioButton.Size = New-Object System.Drawing.Size($Width,$Height)
        $RadioButton.Location = New-Object System.Drawing.Point($DrawX,$DrawY)
        $RadioButton.Checked = $Checked
        $RadioButton.Visible = $true
        If ($Disabled) { $RadioButton.Enabled = $false }
        If ($Text) { $RadioButton.Text = $Text } Else { $RadioButton.Text = $Name }

        $RadioButton
    }

    Catch { Invoke-PetErrorHandler -err $_ }   
}

Write-Debug "......Function New-FormRadioButtonGroup"
Function New-FormRadioButtonGroup {
    [CmdLetBinding()]
    Param(
        [string]$Name,
        [string]$Text,
        [int]$Width,
        [int]$Height,
        [int]$DrawX,
        [int]$DrawY,
        $RadioButtons
    )
    Write-Debug ""
    $timestamp = Get-Date
    Write-Debug $timestamp
    Write-Debug "Execute Function New-FormRadioButtonGroup"
    Try {  
        $RadioButtonGroup = [System.Windows.Forms.GroupBox]::new()
        $RadioButtonGroup.Name = ($Name -Replace '[^a-zA-Z]','') + 'RadioButtonGroup'
        Write-Debug "Build RadioButtonGroup Named:"
        Write-Debug $RadioButtonGroup.Name

        $RadioButtonGroup.Size = New-Object System.Drawing.Size($Width,$Height)
        $RadioButtonGroup.Location = New-Object System.Drawing.Point($DrawX,$DrawY)

        If ($Text) { $RadioButtonGroup.Text = $Text } Else { $RadioButtonGroup.Text = $Name }
        $RadioButtonGroup.Controls.AddRange($RadioButtons)

        $RadioButtonGroup
    }

    Catch { Invoke-PetErrorHandler -err $_ }   
}

Write-Debug "......Function New-FormMaskedTextBox"
Function New-FormMaskedTextBox {
    [CmdLetBinding()]
    [OutputType([System.Windows.Forms.MaskedTextBox ])]
    Param(
        [string]$Name,
        [string]$Text,
        [int]$Index,
        [int]$Width,
        [int]$Height,
        [int]$DrawX,
        [int]$DrawY
    )
    Try {
        $MaskedTextBox = [Windows.Forms.MaskedTextBox]::new()
        $MaskedTextBox.PasswordChar = '*'
        $MaskedTextBox.Name = ($Name -Replace '[^a-zA-Z]','') + 'FormMaskedTextBox'
        $MaskedTextBox.Text = $Text
        $MaskedTextBox.TabIndex = $Index
        $MaskedTextBox.Size = [System.Drawing.Size]::new($Width,$Height)
        $MaskedTextBox.Location = [System.Drawing.Point]::new($DrawX,$DrawY)
        $MaskedTextBox.DataBindings.DefaultDataSourceUpdateMode = 0
        
        $MaskedTextBox
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function New-FormCheckBox"
Function New-FormCheckBox {
    [CmdLetBinding()]
    [OutputType([System.Windows.Forms.Checkbox ])]
    Param(
        [string]$Name,
        [string]$Text,
        [int]$Index,
        [int]$Width,
        [int]$Height,
        [int]$DrawX,
        [int]$DrawY,
        [switch]$checked
    )
    Try {
        $CheckBox = [Windows.Forms.Checkbox]::new()
        $CheckBox.Name = ($Name -Replace '[^a-zA-Z]','') + 'FormCheckBox'
        $CheckBox.Text = $Text
        $CheckBox.TabIndex = $Index
        $CheckBox.Size = [System.Drawing.Size]::new($Width,$Height)
        $CheckBox.Location = [System.Drawing.Point]::new($DrawX,$DrawY)
        $CheckBox.DataBindings.DefaultDataSourceUpdateMode = 0

        If ($checked) { $CheckBox.Checked = $true }
        
        $CheckBox
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function New-FormButton"
Function New-FormButton {
    [CmdLetBinding()]
    [OutputType([System.Windows.Forms.Button])]
    Param(
        [string]$Name,
        [string]$Text,
        [int]$Index,
        [int]$Width,
        [int]$Height,
        [int]$DrawX,
        [int]$DrawY,
        [System.Windows.Forms.AnchorStyles]$Anchor='None',
        [scriptblock]$Action,
        #[string]$CallFunction,
        [switch]$Disabled,
        $buttonIcon
    )
    Try {
        $FormButton = [System.Windows.Forms.Button]::new()
        $FormButton.TabIndex = $Index
        $FormButton.Name = ($Name -Replace '[^a-zA-Z]','') + 'FormButton'
        Write-Debug "Build FormButton Named:"
        Write-Debug $FormButton.Name
        If ($Text) { $FormButton.Text = $Text } Else { $FormButton.Text = $Name}
        $FormButton.Size = [System.Drawing.Size]::new($Width,$Height)
        $FormButton.Location = [System.Drawing.Point]::new($DrawX,$DrawY)
        $FormButton.UseVisualStyleBackColor = $True
        $FormButton.DataBindings.DefaultDataSourceUpdateMode = 0
        If ($Disabled) { $FormButton.Enabled = $false }
        if ($Anchor) { $FormButton.Anchor = $Anchor }
        If ($buttonIcon) { $FormButton.Image = $buttonIcon }
        If ($CallFunction) {
            [string]$script:Command = $CallFunction.toString()
            $FormButton.Add_Click({ Invoke-Expression $Command })
        }
        Else {
            $FormButton.Add_Click($Action)
        }
        $FormButton
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function New-FormDateTimePicker"
Function New-FormDateTimePicker {
    [CmdLetBinding()]
    Param(
        [string]$Name,
        [int]$Width,
        [int]$Height,
        [int]$DrawX,
        [int]$DrawY,
        $Value,
        $MinDate,
        $MaxDate
    )
    Write-Debug ""
    $timestamp = Get-Date
    Write-Debug $timestamp
    Write-Debug "Execute Function New-FormDateTimePicker"
    Try {  
        $DateTimePicker = [System.Windows.Forms.DateTimePicker]::new()
        $DateTimePicker.Name = ($Name -Replace '[^a-zA-Z]','') + 'DateTimePicker'
        Write-Debug "Build DateTimePicker Named:"
        Write-Debug $DateTimePicker.Name

        $DateTimePicker.Size = New-Object System.Drawing.Size($Width,$Height)
        $DateTimePicker.Location = New-Object System.Drawing.Point($DrawX,$DrawY)
        $DateTimePicker.Format = [windows.forms.datetimepickerFormat]::custom
        $DateTimePicker.CustomFormat = "MM/dd/yyyy HH:mm:ss"
        
        If ($Value ) { $DateTimePicker.Value = $Value } Else { $DateTimePicker.Value = $timestamp }
        If ($MinDate) { $DateTimePicker.MinDate = $MinDate }
        If ($MaxDate) { $DateTimePicker.MaxDate = $MaxDate } Else { $DateTimePicker.MaxDate = $timestamp }

        $DateTimePicker
    }

    Catch { Invoke-PetErrorHandler -err $_ }   
}

Write-Debug "......Function New-FormProgressBar"
Function New-FormProgressBar {
    [OutputType([System.Windows.Forms.ProgressBar])]
    [CmdLetBinding()]
    Param(
        [string]$Name,
        [int]$Width,
        [int]$Height,
        [int]$DrawX,
        [int]$DrawY,
        [string]$Style = 'Continuous'
    )
    Write-Debug ""
    $timestamp = Get-Date
    Write-Debug $timestamp
    Write-Debug "Execute Function New-FormProgressBar"
    Try {  
        $ProgressBar = [System.Windows.Forms.ProgressBar]::new()
        $ProgressBar.Name = ($Name -Replace '[^a-zA-Z]','') + 'ProgressBar'
        Write-Debug "Build ProgressBar Named:"
        Write-Debug $ProgressBar.Name
        $ProgressBar.Location = New-Object System.Drawing.Point($DrawX,$DrawY)
        $ProgressBar.Size = New-Object System.Drawing.Size($Width,$Height)
        $ProgressBar.Style = $Style ##  Blocks, Continuous, Marquee
        $ProgressBar.MarqueeAnimationSpeed = 20
        $ProgressBar.Visible = $true

        $ProgressBar
    }

    Catch { Invoke-PetErrorHandler -err $_ }   
}

Write-Debug "......Function Update-FormProgressBar"
Function Update-FormProgressBar {
    
    [CmdLetBinding()]
    Param(
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [int]$Value
    )
    Try {  
        $ProgressBar.Value = $Value
    }

    Catch { Invoke-PetErrorHandler -err $_ }   
}

Write-Debug "......Function New-FormWebView"
Function New-FormWebView {
    [CmdLetBinding()]
    [OutputType([Microsoft.Web.WebView2.WinForms.WebView2])]
    Param(
        [string]$Name,
        [int]$Index,
        [int]$Width,
        [int]$Height,
        [int]$DrawX,
        [int]$DrawY,
        [System.Windows.Forms.AnchorStyles]$Anchor='None',
        [string]$URL,
        [string]$Navigated,
        [string]$DocumentCompleted
    )
    Try {
        $FormWebView = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=$Width;Height=$Height;Url=($Url);AllowNavigation=$true}

        $FormWebView.TabIndex = $Index
        $FormWebView.Name = ($Name -Replace '[^a-zA-Z]','') + 'FormWebView'
        Write-Debug "Build FormWebView Named:"
        Write-Debug $FormWebView.Name
        $FormWebView.ScriptErrorsSuppressed = $true        
        #$FormWebView.Size = [System.Drawing.Size]::new($Width,$Height)
        $FormWebView.Location = [System.Drawing.Point]::new($DrawX,$DrawY)
        $FormWebView.Visible = $true

        If ($Anchor) { $FormWebView.Anchor = $Anchor }
        If ($Navigated) { $FormWebView.Add_Navigated($Navigated) }
        If ($DocumentCompleted) { $FormWebView.Add_DocumentCompleted($DocumentCompleted) }

        $FormWebView
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function New-FormWebView2"
Function New-FormWebView2 {
    [CmdLetBinding()]
    [OutputType([Microsoft.Web.WebView2.WinForms.WebView2])]
    Param(
        [string]$Name,
        [int]$Index,
        [int]$Width,
        [int]$Height,
        [int]$DrawX,
        [int]$DrawY,
        [System.Windows.Forms.AnchorStyles]$Anchor='None',
        [string]$URL
    )
    Try {
        [Microsoft.Web.WebView2.WinForms.WebView2] $FormWebView2 = New-Object 'Microsoft.Web.WebView2.WinForms.WebView2'
        $FormWebView2.CreationProperties = New-Object 'Microsoft.Web.WebView2.WinForms.CoreWebView2CreationProperties'
        $FormWebView2.CreationProperties.UserDataFolder = "$env:USERPROFILE\tmp\wv2\data";
        $FormWebView2.TabIndex = $Index
        $FormWebView2.Name = ($Name -Replace '[^a-zA-Z]','') + 'FormWebView2'
        Write-Debug "Build FormWebView2 Named:"
        Write-Debug $FormWebView2.Name
        #$FormWebView2.Text = $Name
        $FormWebView2.Size = [System.Drawing.Size]::new($Width,$Height)
        $FormWebView2.Location = [System.Drawing.Point]::new($DrawX,$DrawY)
        $FormWebView2.ZoomFactor = 1
        $FormWebView2.Visible = $true

        if ($Anchor) {
            $FormWebView2.Anchor = $Anchor
        }

        $FormWebView2
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function New-DataGridView"
Function New-DataGridView {
    [CmdLetBinding()]
    [OutputType([System.Windows.Forms.DataGridView])]
    Param(
        [string]$Name,
        [int]$Index,
        [int]$Width,
        [int]$Height,
        [int]$DrawX,
        [int]$DrawY,
        [System.Windows.Forms.AnchorStyles]$Anchor='None'
    )

    Try {
        $DataGridView = [System.Windows.Forms.DataGridView]::new()
        $DataGridView.TabIndex = $Index
        $DataGridView.Name = $Name.Replace('[^a-zA-Z]','') + 'DataGridView'
        $DataGridView.AutoSizeColumnsMode = 'AllCells'
        $DataGridView.Size = [System.Drawing.Size]::new($Width,$Height)
        $DataGridView.Location = [System.Drawing.Point]::new($DrawX,$DrawY)
        $DataGridView.DataBindings.DefaultDataSourceUpdateMode = 'OnValidation'

        $DataGridView.DataMember = ""
        if ($Anchor) {
            $DataGridView.Anchor = $Anchor
        }
        $DataGridView
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function Update-DataGridView"
Function Update-DataGridView {
    [CmdLetBinding()]
    [OutputType([System.Windows.Forms.DataGridView])]
    Param(
        [object]$Data,
        [System.Windows.Forms.DataGridView]$DataGridView,
        [hashtable]$RowHighlight
    )
    Write-Debug ""
    $timestamp = Get-Date
    Write-Debug $timestamp
    Write-Debug "Execute FORM Function Update-DataGridView"
    Try {
        #Write-Debug "Data = $Data"
        #If ($Data.GetType().Name -eq 'DataTable') { $DataGridView.DataSource = $Data }
        #Else {
            #$GridData = [System.Collections.ArrayList]::new()
            #$GridData.AddRange(@($Data))
            $GridData = $Data | convertto-datatable
            $DataGridView.DataSource = $GridData
        #}

        if ($RowHighlight) {
            $Cell = $RowHighlight['Cell']
            foreach ($Row in $DataGridView.Rows) {
                [string]$CellValue = $Row.Cells[$Cell].Value
                Write-Verbose ($CellValue.Gettype())
                if ($RowHighlight['Values'].ContainsKey($CellValue)) {
                    Write-Verbose "Setting row based on $Cell cell of $CellValue to $($RowHighlight['Values'][$CellValue]) color"
                    $Row.DefaultCellStyle.BackColor = $RowHighlight['Values'][$CellValue]
                } elseif ($RowHighlight['Values'].ContainsKey('Default')) {
                    Write-Verbose "Setting $Cell cell for $CellValue to $($RowHighlight['Values'].Default) color"
                    $Row.DefaultCellStyle.BackColor = $RowHighlight['Values']['Default']
                }
            }
        }

        Write-Debug 'Exit FORM Function Update-DataGridView with Return:$DataGridView'
        Return $DataGridView
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
        Write-Debug "Exit FORM Function Update-DataGridView with Catch Exception"
    }
}

Write-Debug "......Function New-StatusStrip"
Function New-StatusStrip {
    [CmdLetBinding()]
    [OutputType([System.Windows.Forms.StatusStrip])]
    Param()

    Try {
        $StatusStrip = [System.Windows.Forms.StatusStrip]::new()
        $StatusStrip.Name = 'StatusStrip'
        $StatusStrip.AutoSize = $true
        $StatusStrip.Left = 0
        $StatusStrip.Visible = $true
        $StatusStrip.Enabled = $true
        $StatusStrip.Dock = [System.Windows.Forms.DockStyle]::Bottom
        $StatusStrip.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
        $StatusStrip.LayoutStyle = [System.Windows.Forms.ToolStripLayoutStyle]::Table

        $Operation = [System.Windows.Forms.ToolStripLabel]::new()
        $Operation.Name = 'Operation'
        $Operation.Text = $null
        $Operation.Width = 50
        $Operation.Visible = $true

        $Progress = [System.Windows.Forms.ToolStripLabel]::new()
        $Progress.Name = 'Progress'
        $Progress.Text = $null
        $Progress.Width = 50
        $Progress.Visible = $true

        $ProgressBar = [System.Windows.Forms.ToolStripProgressBar]::new()
        $ProgressBar.Name = 'ProgressBar'
        $ProgressBar.Width = 50
        $ProgressBar.Visible = $false

        $StatusStrip.Items.AddRange(
            [System.Windows.Forms.ToolStripItem[]]@(
                $Operation,
                $Progress,
                $ProgressBar
            )
        )
        $StatusStrip
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function Set-StatusStrip"
Function Set-StatusStrip {
    [CmdLetBinding()]
    [OutputType([System.Windows.Forms.StatusStrip])]
    Param(
        [System.Windows.Forms.StatusStrip]$StatusStrip,
        [string]$Operation = $null,
        [string]$Progress = $null,
        [int]$ProgressBarMinimum = 0,
        [int]$ProgressBarMaximum = 0,
        [int]$ProgressBarValue = 0
    )
    Try {

        if ($null -ne $Operation) {
            $StatusStrip.Items.Find('Operation',$true)[0].Text = $Operation
            $StatusStrip.Items.Find('Operation',$true)[0].Width = 200
            $StatusStrip.Items.Find('Operation',$true)[0].Visible = $true
        }

        if ($null -ne $Progress) {
            $StatusStrip.Items.Find('Progress',$true)[0].Text = $Progress
            $StatusStrip.Items.Find('Progress',$true)[0].Width = 100
            $StatusStrip.Items.Find('Progress',$true)[0].Visible = $true
        }

        if ($null -ne $StatusStrip.Items.Find('ProgressBar',$true)) {
            if ($null -ne $ProgressBarMinimum) {
                $StatusStrip.Items.Find('ProgressBar',$true)[0].Minimum = $ProgressBarMinimum
            }
            if ($null -ne $ProgressBarMaximum) {
                $StatusStrip.Items.Find('ProgressBar',$true)[0].Maximum = $ProgressBarMaximum
            }
            if ($null -ne $ProgressBarValue) {
                $StatusStrip.Items.Find('ProgressBar',$true)[0].Value = $ProgressBarValue
            }
            if ($StatusStrip.Items.Find('ProgressBar',$true)[0].Minimum -eq $StatusStrip.Items.Find('ProgressBar',$true)[0].Maximum ) {
                $StatusStrip.Items.Find('ProgressBar',$true)[0].Visible = $false
            } else {
                $StatusStrip.Items.Find('ProgressBar',$true)[0].Visible = $true
            }
        }
        $StatusStrip
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function Show-WindowsForm"
Function Show-WindowsForm {
    [CmdLetBinding()]
    [OutputType([System.Windows.Forms.Form])]
    Param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Form]$WindowsForm,
        [System.Windows.Forms.Label[]]$FormLabel,
        [System.Windows.Forms.TextBox[]]$FormTextBox,
        [System.Windows.Forms.MaskedTextBox[]]$FormMaskedTextBox,
        [System.Windows.Forms.Button[]]$FormButton,
        [System.Windows.Forms.DataGridView]$DataGridView,
        [System.Windows.Forms.StatusStrip]$StatusStrip,
        [ScriptBlock]$OnLoad,
        [int]$HeaderWidth
    )

    Try {
        if ($PSBoundParameters.Keys -contains 'FormLabel') {
            foreach ($Label in $FormLabel) {
                $WindowsForm.Controls.Add($Label)
            }
        }
        if ($PSBoundParameters.Keys -contains 'FormTextBox') {
            foreach ($TextBox in $FormTextBox) {
                $WindowsForm.Controls.Add($TextBox)
            }
        }
        if ($PSBoundParameters.Keys -contains 'FormMaskedTextBox') {
            foreach ($MaskedTextBox in $FormMaskedTextBox) {
                $WindowsForm.Controls.Add($MaskedTextBox)
            }
        }
        if ($PSBoundParameters.Keys -contains 'FormButton') {
            foreach ($Button in $FormButton) {
                $WindowsForm.Controls.Add($Button)
            }
        }
        if ($PSBoundParameters.Keys -contains 'DataGridView') {
            $WindowsForm.Controls.Add($DataGridView)
        }
        if ($PSBoundParameters.Keys -contains 'StatusStrip') {
            $WindowsForm.Controls.Add($StatusStrip)
        }
        if ($PSBoundParameters.Keys -contains 'OnLoad') {
            $WindowsForm.add_Shown($OnLoad)
        }
        if ($PSBoundParameters.Keys -contains 'HeaderWidth') {
            $WindowsForm.Width = $HeaderWidth + 5
        }

        $WindowsForm.StartPosition = 1

        $WindowsForm
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

Write-Debug "......Function New-MessageBox"
Function New-MessageBox {
    [CmdLetBinding()]
    [OutputType([System.Windows.Forms.Form])]
    Param(
        [string]$text,
        [string]$title,
        [string]$buttons,
        [string]$icon='None'
    )
    Write-Debug ""
    $timestamp = Get-Date
    Write-Debug $timestamp
    Write-Debug "Execute Function New-MessageBox"

    Try {
        If ($buttons -eq '' -OR $null -eq $buttons) { $buttons = 'Ok' } ## OK, OKCancel, YesNoCancel, YesNo ##
        [System.Windows.MessageBox]::Show($text,$title,$buttons,$icon)
        Return 
    }
        Catch { Invoke-PetErrorHandler -err $_ }
}

Write-Debug "......Function New-DropdownMenu"
Function New-DropdownMenu {
    [CmdLetBinding()]
    Param(
        [string]$Name,
        [int]$Index,
        [int]$Width,
        [int]$Height,
        [int]$DrawX,
        [int]$DrawY,
        [array]$MenuOptions,
        [int]$selectedIndex,
        [switch]$readOnly
    )
    Write-Debug ""
    $timestamp = Get-Date
    Write-Debug $timestamp
    Write-Debug "Execute Function New-DropdownMenu"

    Try {
        $DropdownMenu = New-Object system.Windows.Forms.ComboBox
        $DropdownMenu.Name = ($Name -Replace '[^a-zA-Z]','') + 'FormDropDownMenu'
        $DropdownMenu.TabIndex = $Index
        Write-Debug "Build FormButton Named:"
        Write-Debug $DropdownMenu.Name
        $DropdownMenu.Height = $Height
        $DropdownMenu.Width = $Width
        $DropdownMenu.Size = [System.Drawing.Size]::new($Width,$Height)
        $DropdownMenu.Location = [System.Drawing.Point]::new($DrawX,$DrawY)
        $MenuOptions | ForEach-Object {
            Write-Debug "Adding Menu Option: $_"
            [void] $DropdownMenu.Items.Add($_)
        }
        # Select the default value
        If ($selectedIndex) { $DropdownMenu.SelectedIndex = $selectedIndex } Else { $DropdownMenu.SelectedIndex = 0 }
        If ($readOnly) { $DropdownMenu.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList; }

        Return $DropdownMenu
    }
    Catch {
        $PSCmdlet.ThrowTerminatingError($_)
        Return $null
    }
}

Write-Debug "......Function ConvertTo-DataTable"
function ConvertTo-DataTable {
    <#
    .Synopsis
        Creates a DataTable from an object
    .Description
        Creates a DataTable from an object, containing all properties (except built-in properties from a database)
    .Example
        Get-ChildItem| Select Name, LastWriteTime | ConvertTo-DataTable
    .Link
        Select-DataTable
    .Link
        Import-DataTable
    .Link
        Export-Datatable
    #> 
    [OutputType([Data.DataTable])]
    param(
    # The input objects
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)]
    [PSObject[]]
    $InputObject
    ) 
    

    begin { 
        Write-Debug ""
        $timestamp = Get-Date
        Write-Debug $timestamp
        Write-Debug "Execute Function ConvertTo-DataTable"
        $outputDataTable = new-object Data.datatable   
        #$knownColumns = @{}
    } 

    process {         
        foreach ($In in $InputObject) { 
            $DataRow = $outputDataTable.NewRow()   
            $isDataRow = $in.psobject.TypeNames -like "*.DataRow*" -as [bool]

            $simpleTypes = ('System.Boolean', 'System.Byte[]', 'System.Byte', 'System.Char', 'System.Datetime', 'System.Decimal', 'System.Double', 'System.Guid', 'System.Int16', 'System.Int32', 'System.Int64', 'System.Single', 'System.UInt16', 'System.UInt32', 'System.UInt64')

            $SimpletypeLookup = @{}
            foreach ($s in $simpleTypes) {
                $SimpletypeLookup[$s] = $s
            }            
            
            foreach($property in $In.PsObject.properties) {   
                if ($isDataRow -and 'RowError', 'RowState', 'Table', 'ItemArray', 'HasErrors' -contains $property.Name) { continue }
                $propName = $property.Name
                $propValue = $property.Value
                $IsSimpleType = $SimpletypeLookup.ContainsKey($property.TypeNameOfValue)

                if (-not $outputDataTable.Columns.Contains($propName)) {   
                    $outputDataTable.Columns.Add((
                        New-Object Data.DataColumn -Property @{
                            ColumnName = $propName
                            DataType = if ($issimpleType) {
                                $property.TypeNameOfValue
                            } else {
                                'System.Object'
                            }
                        }
                    ))
                }                   
                
                $DataRow.Item($propName) = if ($isSimpleType -and $propValue) {
                    $propValue
                } elseif ($propValue) {
                    [PSObject]$propValue
                } else {
                    [DBNull]::Value
                }
            }   
            $outputDataTable.Rows.Add($DataRow)   
        } 
    }  
      
    end 
    {
        Write-Debug 'Exit Function ConvertTo-DataTable with Return $outputDataTable'
        ,$outputDataTable
    }
}

Write-Debug "......Function ConvertFrom-HTMLTable"
function ConvertFrom-HTMLTable {
    <#
    .SYNOPSIS
    Function for converting ComObject HTML object to common PowerShell object.

    .DESCRIPTION
    Function for converting ComObject HTML object to common PowerShell object.
    ComObject can be retrieved by (Invoke-WebRequest).parsedHtml or IHTMLDocument2_write methods.

    In case table is missing column names and number of columns is:
    - 2
        - Value in the first column will be used as object property 'Name'. Value in the second column will be therefore 'Value' of such property.
    - more than 2
        - Column names will be numbers starting from 1.

    .PARAMETER table
    ComObject representing HTML table.

    .PARAMETER tableName
    (optional) Name of the table.
    Will be added as TableName property to new PowerShell object.

    .EXAMPLE
    $pageContent = Invoke-WebRequest -Method GET -Headers $Headers -Uri "https://docs.microsoft.com/en-us/mem/configmgr/core/plan-design/hierarchy/log-files"
    $table = $pageContent.ParsedHtml.getElementsByTagName('table')[0]
    $tableContent = @(ConvertFrom-HTMLTable $table)

    Will receive web page content >> filter out first table on that page >> convert it to PSObject

    .EXAMPLE
    $Source = Get-Content "C:\Users\Public\Documents\MDMDiagnostics\MDMDiagReport.html" -Raw
    $HTML = New-Object -Com "HTMLFile"
    $HTML.IHTMLDocument2_write($Source)
    $HTML.body.getElementsByTagName('table') | % {
        ConvertFrom-HTMLTable $_
    }

    Will get web page content from stored html file >> filter out all html tables from that page >> convert them to PSObjects

    .Link
    Ondrej Sebela:
    https://doitpsway.com/extracting-html-table-from-a-web-page-or-html-file-and-converting-it-into-powershell-object
    https://github.com/ztrhgf/useful_powershell_functions/blob/master/ConvertFrom-HTMLTable.ps1
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.__ComObject] $table,

        [string] $tableName
    )

    $twoColumnsWithoutName = 0

    if ($tableName) { $tableNameTxt = "'$tableName'" }

    $columnName = $table.getElementsByTagName("th") | ForEach-Object { $_.innerText -replace "^\s*|\s*$" }

    if (!$columnName) {
        $numberOfColumns = @($table.getElementsByTagName("tr")[0].getElementsByTagName("td")).count
        if ($numberOfColumns -eq 2) {
            ++$twoColumnsWithoutName
            Write-Verbose "Table $tableNameTxt has two columns without column names. Resultant object will use first column as objects property 'Name' and second as 'Value'"
        } elseif ($numberOfColumns) {
            Write-Warning "Table $tableNameTxt doesn't contain column names, numbers will be used instead"
            $columnName = 1..$numberOfColumns
        } else {
            throw "Table $tableNameTxt doesn't contain column names and summarization of columns failed"
        }
    }

    if ($twoColumnsWithoutName) {
        # table has two columns without names
        $property = [ordered]@{ }

        $table.getElementsByTagName("tr") | ForEach-Object {
            # read table per row and return object
            $columnValue = $_.getElementsByTagName("td") | ForEach-Object { $_.innerText -replace "^\s*|\s*$" }
            if ($columnValue) {
                # use first column value as object property 'Name' and second as a 'Value'
                $property.($columnValue[0]) = $columnValue[1]
            } else {
                # row doesn't contain <td>
            }
        }
        if ($tableName) {
            $property.TableName = $tableName
        }

        New-Object -TypeName PSObject -Property $property
    } else {
        # table doesn't have two columns or they are named
        $table.getElementsByTagName("tr") | ForEach-Object {
            # read table per row and return object
            $columnValue = $_.getElementsByTagName("td") | ForEach-Object { $_.innerText -replace "^\s*|\s*$" }
            if ($columnValue) {
                $property = [ordered]@{ }
                $i = 0
                $columnName | ForEach-Object {
                    $property.$_ = $columnValue[$i]
                    ++$i
                }
                if ($tableName) {
                    $property.TableName = $tableName
                }

                New-Object -TypeName PSObject -Property $property
            } else {
                # row doesn't contain <td>, its probably row with column names
            }
        }
    }
}

Write-Debug "......Function New-ImageList"
Function New-ImageList {
    [CmdLetBinding()]
    Param(
        [int]$ImageWidth,
        [int]$ImageHeight,
        [string]$ColorDepth = "Depth32Bit",
        $Images
    )
    Write-Debug ""
    $timestamp = Get-Date
    Write-Debug $timestamp
    Write-Debug "Execute Function New-ImageList"

    $ImageList = New-Object System.Windows.Forms.ImageList
    $ImageList.ImageSize = [System.Drawing.Size]::new($ImageWidth,$ImageHeight)
    $ImageList.ColorDepth = $ColorDepth
    
    $i = 0
    ForEach ($image in $Images) {
        $ImageList.Images.Add($i,$image) | out-Null
        $i++
    }

    Write-Debug "Exit Function New-ImageList"
    $ImageList
}

Write-Debug "......Function New-ImageListView"
Function New-ImageListView {
    [OutputType([System.Windows.Forms.ListView])]
    [CmdLetBinding()]
    Param(
        [string]$Name,
        [int]$Width,
        [int]$Height,
        [int]$DrawX,
        [int]$DrawY,
        [bool]$MultiSelect = $false,
        [bool]$LabelWrap = $false,
        [string]$View,
        [array]$Columns,
        [int]$ColumnWidth,
        [string]$BackColor,
        $ImageList,
        $form
    )
    Write-Debug ""
    $timestamp = Get-Date
    Write-Debug $timestamp
    Write-Debug "Execute Function New-ImageListView"
    Try {
        $ImageListView = New-Object System.Windows.Forms.ListView
        $ImageListView.Name = ($Name -Replace '[^a-zA-Z]','') + 'ImageListView'
        Write-Debug "Build ImageListView Named:"
        Write-Debug $ImageListView.Name
        $ImageListView.Size = New-Object System.Drawing.Size($Width,$Height)
        $ImageListView.Location = New-Object System.Drawing.Point($DrawX,$DrawY)
        $ImageListView.Visible = $True
        $ImageListView.MultiSelect = $MultiSelect
        $ImageListView.LabelWrap = $LabelWrap
        $ImageListView.LargeImageList = $ImageList

        If ($Columns) { Foreach ($column in $columns) { $ImageListView.Columns.Add($column,$ColumnWidth) | Out-Null } } Else { $ImageListView.Columns.Add('',250) | Out-Null }
        If ($BackColor) { $ImageListView.BackColor = $BackColor }
    }
    Catch { Invoke-PetErrorHandler -err $_ }

    Write-Debug "Exit Function New-ImageListView"
    Return $ImageListView
}

Write-Debug "......Function Show-SaveFileGUI"
Function Show-SaveFileGUI {
    Write-Debug ""
    $timestamp = Get-Date
    Write-Debug $timestamp
    Write-Debug "Execute Function Save-FileGUI"
    $filePath = Show-SaveFileDialogGUI("Save As",'*.*')
    #Remove-Variable filecontents
    #Remove-Variable filePath
    Write-Debug "Exit Function Function Save-FileGUI with Return"
    Return $filePath
} 

Write-Debug "......Function Show-OpenFileGUI"
Function Show-OpenFileGUI() {
    Write-Debug ""
    $timestamp = Get-Date
    Write-Debug $timestamp
    Write-Debug "Execute Function Open-FileGUI"
    $fileContents = Get-OpenFileDialogGUI("Open",'*.*')
    Write-Debug "Exit Function Function Open-FileGUI with Return fileContects"
    Return $fileContents
} 

Write-Debug "......Function Set-FileDialogFilterIndex"
Function Set-FileDialogFilterIndex ($filter) {
    $index = 7
    If ($filter -like "bak") { $index = 1 }
    If ($filter -like "crt") { $index = 2 }
    If ($filter -like "key") { $index = 3 }
    If ($filter -like "pfx") { $index = 4 }
    If ($filter -like "reg") { $index = 5 }
    If ($filter -like "txt") { $index = 6 }
    Write-Debug "index = $index"
    Return $index
} 

Write-Debug "......Function Set-FileDialogFilterList"
Function Set-FileDialogFilterList {
    Return "BAK (*.bak)|*.bak|CRT (*.crt)|*.crt|KEY (*.key)|*.key|PFX (*.pfx)|*.pfx|REG (*.reg)|*.reg|TXT (*.txt)|*.txt|All Files (*.*)|*.*"
} 

Write-Debug "......Function Show-SaveFileDialogGUI"
Function Show-SaveFileDialogGUI ([string]$formTitle,[string]$formFilter) {
    Write-Debug "formTitle = $formTitle"
    Write-Debug "formFilter = $formFilter"
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Title = $formTitle
    $SaveFileDialog.initialDirectory = $env:HOMEPATH
    $SaveFileDialog.filter = Get-FileDialogFilterList
    $SaveFileDialog.filterIndex = Get-FileDialogFilterIndex($formFilter)
    $SaveFileDialog.ShowDialog() | Out-Null
    $SaveFileDialog.filename
    $fullFilePath = $SaveFileDialog.filename
    Remove-Variable SaveFileDialog
    Return $fullFilePath
}

Write-Debug "......Function Show-OpenFileDialogGUI"
Function Show-OpenFileDialogGUI ([string]$formTitle,[string]$formFilter) {
    Write-Debug "formTitle = $formTitle"
    Write-Debug "formFilter = $formFilter"
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = $formTitle
    $OpenFileDialog.initialDirectory = $env:HOMEPATH
    $OpenFileDialog.filter = Get-FileDialogFilterList
    $OpenFileDialog.FilterIndex = Get-FileDialogFilterIndex($formFilter)
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
    $fullFilePath = $OpenFileDialog.filename
    Remove-Variable OpenFileDialog
    Return $fullFilePath
}