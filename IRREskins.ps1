Clear-Host

<#
    IRREskins
    Script de Mise à jour des skins BOS
    par -IRRE-Biluf / biluf@carlton2001.com
	
	TEST UTF8
	
    Prérequis :
    - Powershell v3.0+
    - 7-Zip
#>

#region XAML

[xml]$XAML  = @'

<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
Title="IRREskins" FontFamily="Calibri" FontSize="14" Width="600" SizeToContent="Height" ResizeMode="NoResize">
<Grid Margin="10,10,10,10">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition />
        </Grid.ColumnDefinitions>
        <DockPanel Grid.Row="0" Margin="0">
            <Label Content="IRREskins" FontSize="22" FontWeight="Bold"/>
            <Label x:Name="LB_Version" Content="v0.0" HorizontalContentAlignment="Right" Opacity="0.5"/>
        </DockPanel>
        <GroupBox x:Name="GP_Paths" Header="Chemins d'installations" Grid.Row="1" Margin="0,10,0,0" Padding="5">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="27"/>
                    <RowDefinition Height="5"/>
                    <RowDefinition Height="27"/>
                    <RowDefinition Height="5"/>
                    <RowDefinition Height="27"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="5"/>
                    <ColumnDefinition />
                    <ColumnDefinition Width="5"/>
                    <ColumnDefinition Width="30"/>
                </Grid.ColumnDefinitions>
                <Label Content="StandAlone" Grid.Row="0" Grid.Column="0"/>
                <TextBox x:Name="TB_StandAlone" Grid.Row="0" Grid.Column="2" VerticalContentAlignment="Center" IsReadOnly="True" MaxLines="1"/>
                <Button x:Name="BT_StandAlone" Content="..." Grid.Row="0" Grid.Column="4" ToolTip="Chemin d'installation de BOS StandAlone (non Steam)"/>
                <Label Content="Steam" Grid.Row="2" Grid.Column="0"/>
                <TextBox x:Name="TB_Steam" Grid.Row="2" Grid.Column="2" VerticalContentAlignment="Center" IsReadOnly="True" MaxLines="1"/>
                <Button x:Name="BT_Steam" Content="..." Grid.Row="2" Grid.Column="4" ToolTip="Chemin d'installation de Steam"/>
                <Label Content="7-Zip" Grid.Row="4" Grid.Column="0"/>
                <TextBox x:Name="TB_SZip" Grid.Row="4" Grid.Column="2" VerticalContentAlignment="Center" IsReadOnly="True" MaxLines="1"/>
                <Button x:Name="BT_SZip" Content="..." Grid.Row="4" Grid.Column="4" ToolTip="Chemin d'installation 7-Zip"/>
            </Grid>
        </GroupBox>
        <GroupBox x:Name="GP_Preferences" Header="Préférences" Grid.Row="2" Margin="0,10,0,0" Padding="5">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="27"/>
                    <RowDefinition Height="5"/>
                    <RowDefinition Height="27"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="5"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="10"/>
                    <ColumnDefinition Width="30"/>
                    <ColumnDefinition Width="60"/>
                    <ColumnDefinition />
                </Grid.ColumnDefinitions>
                <Label Content="Mises à jour :" Grid.Row="0" Grid.Column="0"/>
                <DockPanel Grid.Row="0" Grid.Column="2">
                    <CheckBox x:Name="CB_StandAlone" VerticalContentAlignment="Center" Content="StandAlone" Margin="0,0,10,0"/>
                    <CheckBox x:Name="CB_Steam" VerticalContentAlignment="Center" Content="Steam"/>
                </DockPanel>
                <CheckBox x:Name="CB_KeepFiles" VerticalContentAlignment="Center" Content="Conserver les 7-Zips téléchargés" Grid.Row="2" Grid.Column="2"/>
                <Button x:Name="BT_OpenFilesFolder" Content="..." Width="30" Grid.Row="2" Grid.Column="4" ToolTip="Ouvrir le dossier contenant les fichiers"/>
                <Label Content="Qualité des skins" Grid.Row="0" Grid.Column="6" HorizontalAlignment="Center" Margin="0,0,10,0"/>
                <Grid Grid.Row="2" Grid.Column="6" ToolTip="2K = 5,4Mo par skin, 4K = 21,8Mo par skin">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition />
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <Label Content="2K" Grid.Column="0"/>
                    <Slider x:Name="SL_Quality" Value="1" Maximum="1" Grid.Column="1" VerticalAlignment="Center" Interval="1" IsSnapToTickEnabled="True"/>
                    <Label Content="4K" Grid.Column="2"/>
                </Grid>
            </Grid>
        </GroupBox>
        <GroupBox x:Name="GB_Execution" Header="Execution" Grid.Row="3" Margin="0,10,0,0" Padding="5" Visibility="Visible">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition />
                    <RowDefinition Height="2"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="27"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="100" />
                    <ColumnDefinition />
                    <ColumnDefinition Width="100" />
                    <ColumnDefinition Width="5" />
                    <ColumnDefinition Width="100" />
                </Grid.ColumnDefinitions>
                <RichTextBox x:Name="RTB_Execution" Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="5" HorizontalAlignment="Left" Height="300" Margin="0,0,0,0" IsReadOnly="True" VerticalScrollBarVisibility="Auto">
                    <FlowDocument>
                        <Paragraph x:Name="P_Execution"/>
                    </FlowDocument>
                </RichTextBox>
                <ProgressBar x:Name="PB_Progress" Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="5" Background="{x:Null}" BorderBrush="{x:Null}"/>
                <Button x:Name="BT_Scan" Content="Scanner" Grid.Row="3" Grid.Column="2"/>
                <Button x:Name="BT_Apply" Content="Mettre à jour" Grid.Row="3" Grid.Column="4"/>
                <Button x:Name="BT_Quit" Content="Fermer" Grid.Row="3" Grid.Column="0"/>
            </Grid>
        </GroupBox>
        <GroupBox x:Name="GB_Prerequis_SZip" Header="Prérequis manquant" Grid.Row="3" Margin="0,10,0,0" Padding="5" Visibility="Collapsed">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>
                <TextBox Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="2" BorderBrush="{x:Null}" SelectionBrush="{x:Null}" IsReadOnly="True" Background="{x:Null}" TextWrapping="Wrap" BorderThickness="0" Foreground="Red">
                    ATTENTION ! L'installation de 7-Zip est un prérequis pour le fonctionnement de ce script, indiquez le bon chemin d'installation ou si ce n'est pas déjà fait, téléchargez le : 
                </TextBox>
                <Button x:Name="BT_LinkSZip" Content="http://www.7-zip.org"  Grid.Row="1" Grid.Column="0" Margin="0,10,0,0" Padding="5" />
            </Grid>
        </GroupBox>
        <GroupBox x:Name="GB_Prerequis_GamePath" Header="Prérequis manquant" Grid.Row="3" Margin="0,10,0,0" Padding="5" Visibility="Collapsed">
            <Grid>
                <TextBox BorderBrush="{x:Null}" SelectionBrush="{x:Null}" IsReadOnly="True" Background="{x:Null}" TextWrapping="Wrap" BorderThickness="0" Foreground="Red">
                    ATTENTION ! Vous devez paramétrer au moins un chemin d'installation de BOS valide pour exécuter le script.
                </TextBox>
            </Grid>
        </GroupBox>
    </Grid>
</Grid>
</Window>

'@

#endregion

#region Functions

    # GUI - Ecriture dans une RTB
    function New_Line { # Formatage d'une nouvelle ligne RTB
        param (
            [string]$Text,
            [string]$Foreground,
            [string]$Background,
            [string]$Paragraph,
            [switch]$Bold,
            [switch]$Italic,
            [switch]$Break,
            [switch]$ScrollOff
        )
        switch ($Paragraph) {
            "P_Execution"		{ $PAR = $gui.P_Execution; $RTB = $gui.RTB_Execution; break }
        }
        $Run = New-Object System.Windows.Documents.Run
        if ($PSBoundParameters['Break']) { $PAR.Inlines.Add((New-Object System.Windows.Documents.LineBreak)) }
        else {
            if ($PSBoundParameters['Foreground'])	{ $Run.Foreground = $Foreground }
            if ($PSBoundParameters['Background'])	{ $Run.Background = $Background }
            if ($PSBoundParameters['Bold'])			{ $Run.FontWeight = 'Bold' }
            if ($PSBoundParameters['Italic'])		{ $Run.FontStyle = 'Italic' }
            $Run.Text = $Text
            $PAR.Inlines.Add($Run)
            $PAR.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
        }
        if (!($PSBoundParameters['ScrollOff'])) {
            $RTB.ScrollToEnd()
        }
    }
    # GUI - Popup de sélection d'un dossier
    function Select_Folder ($Message) {
        $Shell = new-object -com Shell.Application
        $objFolder = $Shell.BrowseForFolder(0, $Message, 0, 17)
        return $objFolder.Self.Path
    }
    # GUI - Affichage des résultats du Scan
    function Display_ScanResults ($Database) {
        $Title = "installation $Database".ToUpper()
        $ToDeleteDatabase = "ToDelete$Database"
        $SkinsToUpdateDatabase = "SkinsToUpdate$Database"
        New_Line -Text $Title -Bold -Paragraph "P_Execution"
        New_Line -Break -Paragraph "P_Execution"
        if ($ScriptHT.$ToDeleteDatabase -gt 0) {
            New_Line -Text "Skins obsolètes à supprimer :" -Bold -Foreground "Red" -Paragraph "P_Execution"
            New_Line -Break -Paragraph "P_Execution"
            for ($i = 0; $i -lt $ScriptHT.$ToDeleteDatabase.Count; $i++) {
                New_Line -Text "$($ScriptHT.$ToDeleteDatabase[$i])" -Paragraph "P_Execution"
            }
            New_Line -Break -Paragraph "P_Execution"
        }
        New_Line -Text "Collections à installer / Mettre à jour:" -Bold -Foreground "Green" -Paragraph "P_Execution"
        New_Line -Break -Paragraph "P_Execution"
        if ($dtsHT.$SkinsToUpdateDatabase.Item.Count -eq 1) {
            $FileDate = $dtsHT.$SkinsToUpdateDatabase.FileDate
            $FileDate = ("{0}/{1}/{2}" -f $FileDate.Substring(6,2),$FileDate.Substring(4,2),$FileDate.Substring(0,4))
            $Collection = ("{0}`t{1}`tQualité {2} - {3}`t{4}"-f    $dtsHT.$SkinsToUpdateDatabase.Tag,
                                                $dtsHT.$SkinsToUpdateDatabase.SkinCollection,
                                                $dtsHT.$SkinsToUpdateDatabase.Quality,
                                                $dtsHT.$SkinsToUpdateDatabase.FileSize,
                                                $FileDate)
            New_Line -Text $Collection -Paragraph "P_Execution"
        }
        elseif ($dtsHT.$SkinsToUpdateDatabase.Item.Count -gt 1) {
            for ($i = 0; $i -lt $dtsHT.$SkinsToUpdateDatabase.Item.Count; $i++) {
                $FileDate = $dtsHT.$SkinsToUpdateDatabase.FileDate[$i]
                $FileDate = ("{0}/{1}/{2}" -f $FileDate.Substring(6,2),$FileDate.Substring(4,2),$FileDate.Substring(0,4))
                $Collection = ("{0}`t{1}`tQualité {2} - {3}`t{4}"-f    $dtsHT.$SkinsToUpdateDatabase.Tag[$i],
                                                    $dtsHT.$SkinsToUpdateDatabase.SkinCollection[$i],
                                                    $dtsHT.$SkinsToUpdateDatabase.Quality[$i],
                                                    $dtsHT.$SkinsToUpdateDatabase.FileSize[$i],
                                                    $FileDate)
                New_Line -Text $Collection -Paragraph "P_Execution"
            }
        }
        else {
            New_Line -Text "   \o/ Collections à jour" -Foreground "Green" -Paragraph "P_Execution"
            $gui.GP_Paths.IsEnabled = $true
            $gui.GP_Preferences.IsEnabled = $true
            $gui.BT_Scan.IsEnabled = $true
            $gui.BT_Quit.IsEnabled = $true
        }
    }
    # SCRIPT - Mise à jour du Script
    function Maj_Script ($Version,$InstallFile) {
        # Ajoute les fonctions au Runspace
        $FunctionList = @("Download_Files","Unzip_File","New_Line")
        $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($Function in $FunctionList) {
            $FunctionDefinition = Get-Content Function:\$Function
            $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $Function, $FunctionDefinition
            $initialSessionState.Commands.Add($SessionStateFunction)
        }
        # Création du Runspace
        $Runspace = [runspacefactory]::CreateRunspace($initialSessionState)
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions	 = "ReuseThread"
        $Runspace.Open()
        $Runspace.SessionStateProxy.SetVariable("Version",$Version)
        $Runspace.SessionStateProxy.SetVariable("InstallFile",$InstallFile)
        $Runspace.SessionStateProxy.SetVariable("gui",$gui)
        $Runspace.SessionStateProxy.SetVariable("ScriptHT",$ScriptHT)
        # Code d'exécution du Runspace
        $Code = {
            if (Test-Path $ScriptHT.Config.SZip) {
                $gui.Window.Dispatcher.invoke("Normal",[action]{
                    $gui.BT_Scan.Visibility = "Collapsed"
                    $gui.BT_Apply.Visibility = "Collapsed"
                    $gui.GB_Execution.Visibility = "Visible"
                    $gui.GB_Prerequis_GamePath.Visibility = "Collapsed"
                    $gui.GB_Prerequis_SZip.Visibility = "Collapsed"
                    New_Line -Text "Mise à jour du script disponible en version $Version" -Bold -Paragraph "P_Execution"
                })
                $ZipFile = "$($ScriptHT.DwlFolder)\$($InstallFile.Split('/')[$InstallFile.Split('/').Count-1])"
                $OnlineMaj = $InstallFile
                Download_Files -Url $OnlineMaj -TargetFile $ZipFile
                $gui.Window.Dispatcher.invoke("Normal",[action]{ New_Line -Text "Installation effectuée" -Paragraph "P_Execution" })            
                Unzip_File -ZipFile $ZipFile -DestFolder $ScriptHT.ScriptPath
                Remove-Item -Path $ZipFile -Force
                $gui.Window.Dispatcher.invoke("Normal",[action]{ New_Line -Text "Veuillez redémarrer le script" -Bold -Foreground "Green" -Paragraph "P_Execution" })
            }
        }
        # Invocation du Runspace
        $PSinstance	 = [powershell]::Create().AddScript($Code)
        $PSinstance.Runspace = $Runspace
        $PSinstance.BeginInvoke()
    }
     # SCRIPT - Vérification du Path de BOS StandAlone
    function Check_Path_StandAlone ($Path) {
        $ScriptHT.Config.StandAlonePath = $null
        # Je vérifie si il-2.exe est bien présent (si Path non null)
        if ($Path -ne $null) {
            $checkPath = $Path.ToLower().Replace("\data\graphics\skins","\bin\game\il-2.exe")
            # Si oui alors validation partout
            if (Test-Path $checkPath) { $ScriptHT.Config.StandAlonePath = $checkPath.Replace("\bin\game\il-2.exe","\data\graphics\skins") }
        }
        # Si non alors recherche regedit
        if (($Path -eq $null) -or ($ScriptHT.Config.StandAlonePath -eq $null)) {
            (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\FirewallRules').PSObject.Members.Name | ForEach-Object {
                if ($_ -like "*il-2.exe") {
                    $checkPath = $_.Split('}')[1].ToLower()
                    if (Test-Path $checkPath) { $ScriptHT.Config.StandAlonePath = $checkPath.Replace("\bin\game\il-2.exe","\data\graphics\skins") }
                }
            }
        }
    }
    # SCRIPT - Vérification du Path de Steam
    function Check_Path_Steam ($Path) {
        $ScriptHT.Config.SteamPath = $null
        # Je vérifie si il-2.exe est bien présent (si Path non null)
        if ($Path -ne $null) {
            $checkPath = $Path.ToLower().Replace("\steamapps\common\il-2 sturmovik battle of stalingrad\data\graphics\skins","\steamApps\common\il-2 sturmovik battle of stalingrad\bin\game\il-2.exe")
            # Si oui alors validation partout
            if (Test-Path $checkPath) { $ScriptHT.Config.SteamPath = $checkPath.Replace("\bin\game\il-2.exe","\data\graphics\skins") }
        }
        # Si non alors recherche regedit
        if (($Path -eq $null) -or ($ScriptHT.Config.SteamPath -eq $null)) {
            if (Test-Path 'HKLM:\SOFTWARE\Valve\Steam') { $checkPath = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Valve\Steam').InstallPath }
            if (Test-Path 'HKLM:\SOFTWARE\Wow6432Node\Valve\Steam') { $checkPath = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\Valve\Steam').InstallPath }
            if ($checkPath) {
                $checkPath = ($checkPath + "\steamApps\common\il-2 sturmovik battle of stalingrad\bin\game\il-2.exe").ToLower()
                # Si ok alors validation partout
                if (Test-Path $checkPath) { $ScriptHT.Config.SteamPath = $checkPath.Replace("\bin\game\il-2.exe","\data\graphics\skins") }
            }
        }
    }
    # SCRIPT - Vérification du Path de 7-Zip
    function Check_Path_SZip ($Path) {
        $ScriptHT.Config.SZip = $null
        # Je vérifie que 7-Zip est bien présent (si Path non null) et fini par l'exe
        if (($Path -ne $null) -and (Test-Path $Path) -and ($Path.EndsWith("7z.exe"))) {
            $ScriptHT.Config.SZip = $Path.ToLower()
        }
        # Sinon je cherche
        else {
            # Si on fournit un folder
            if ($Path -ne $null) {
                # Si oui alors validation partout
                $testPath = "$Path\7z.exe"
                if (Test-Path $testPath) { $ScriptHT.Config.SZip = $testPath.ToLower() }
            }
            # Si non alors recherche dans les ProgramFiles
            if (($Path -eq $null) -or ($ScriptHT.Config.SZip -eq $null)) {
                $Paths = @(${env:ProgramFiles(x86)},$env:ProgramFiles)
                foreach ($Path in $Paths) {
                    $testPath = "$Path\7-Zip\7z.exe"
                    if (Test-Path $testPath) { $ScriptHT.Config.SZip =  $testPath.ToLower() }
                }
            }
        }
    }
    # SCRIPT - Application de la Config
    function Apply_Config {
        # StandAlone
        if ($ScriptHT.Config.StandAlonePath) {
            $gui.TB_StandAlone.Text = $ScriptHT.Config.StandAlonePath
            $gui.TB_StandAlone.Background = "LightGreen"
            $gui.TB_StandAlone.IsEnabled = $true
            $gui.CB_StandAlone.IsEnabled = $true
            if ($ScriptHT.Config.Pref_StandAlone -ne $null) {
                $gui.CB_StandAlone.IsChecked = $ScriptHT.Config.Pref_StandAlone
            }
        }
        else {
            $ScriptHT.Config.Pref_StandAlone = $false
            $gui.TB_StandAlone.Text = "Chemin invalide"
            $gui.TB_StandAlone.Background = "DarkGray"
            $gui.TB_StandAlone.IsEnabled = $false
            $gui.CB_StandAlone.IsChecked = $false
            $gui.CB_StandAlone.IsEnabled = $false
        }
        # Steam
        if ($ScriptHT.Config.SteamPath) {
            $gui.TB_Steam.Text = $ScriptHT.Config.SteamPath
            $gui.TB_Steam.Background = "LightGreen"
            $gui.TB_Steam.IsEnabled = $true
            $gui.CB_Steam.IsEnabled = $true
            if ($ScriptHT.Config.Pref_Steam -ne $null) {
                $gui.CB_Steam.IsChecked = $ScriptHT.Config.Pref_Steam
            }
        }
        else {
            $ScriptHT.Config.Pref_Steam = $false
            $gui.TB_Steam.Text = "Chemin invalide"
            $gui.TB_Steam.Background = "DarkGray"
            $gui.TB_Steam.IsEnabled = $false
            $gui.CB_Steam.IsChecked = $false
            $gui.CB_Steam.IsEnabled = $false
        }
        # Gestion des conflits Checkboxs
        if (($ScriptHT.Config.StandAlonePath) -and ($ScriptHT.Config.SteamPath -eq $null)) {
            $gui.CB_StandAlone.IsEnabled = $true
            $gui.CB_StandAlone.IsChecked = $true
        }
        if (($ScriptHT.Config.SteamPath) -and ($ScriptHT.Config.StandAlonePath -eq $null)) {
            $gui.CB_Steam.IsEnabled = $true
            $gui.CB_Steam.IsChecked = $true
        }
        # Si les deux prefs sont null et qu'un au moins des patchs est valide, cocher le path valide
        if (($ScriptHT.Config.Pref_StandAlone -ne $true) -and ($ScriptHT.Config.Pref_Steam -ne $true)) {
            if ($ScriptHT.Config.StandAlonePath) {
                $gui.CB_StandAlone.IsChecked = $true
                $ScriptHT.Config.Pref_StandAlone = $true
            }
            if ($ScriptHT.Config.SteamPath) {
                $gui.CB_Steam.IsChecked = $true
                $ScriptHT.Config.Pref_Steam = $true
            }
        }
        # SZip
        if ($ScriptHT.Config.SZip) {
            $gui.TB_SZip.Text = $ScriptHT.Config.SZip
            $gui.TB_SZip.Background = "LightGreen"
            $gui.TB_SZip.IsEnabled = $true
        }
        else {
            $gui.TB_SZip.Text = "Chemin invalide"
            $gui.TB_SZip.Background = "DarkGray"
            $gui.TB_SZip.IsEnabled = $false       
        }
        # Gestion des GroupBox
        if (!($ScriptHT.Config.SZip)) {
            $gui.GB_Execution.Visibility = "Collapsed"
            $gui.GB_Prerequis_GamePath.Visibility = "Collapsed"
            $gui.GB_Prerequis_SZip.Visibility = "Visible"
        }
        elseif (($ScriptHT.Config.StandAlonePath -eq $null) -and ($ScriptHT.Config.SteamPath -eq $null)) {
            $gui.GB_Execution.Visibility = "Collapsed"
            $gui.GB_Prerequis_GamePath.Visibility = "Visible"
            $gui.GB_Prerequis_SZip.Visibility = "Collapsed"
        }
        else {
            $gui.GB_Execution.Visibility = "Visible"
            $gui.GB_Prerequis_GamePath.Visibility = "Collapsed"
            $gui.GB_Prerequis_SZip.Visibility = "Collapsed"
        }
        # KeepFiles
        if ($ScriptHT.Config.Pref_KeepFiles -ne $null) {
            $gui.CB_KeepFiles.IsChecked = $ScriptHT.Config.Pref_KeepFiles
        }
        else { $gui.CB_KeepFiles.IsChecked = $true }
        # Quality
        switch ($ScriptHT.Config.Pref_Quality) {
            "2K" { $gui.SL_Quality.Value = 0; break }
            "4K" { $gui.SL_Quality.Value = 1; break }
        }
        Save_Config
    }
    # SCRIPT - Sauvegarde de la config en local
    function Save_Config {
        New-Item -Path $ScriptHT.ConfigFile -ItemType File -Value $($ScriptHT.Config | ConvertTo-Json) -Force | Out-Null
    }
    # SCRIPT - Téléchargement des Zips
    function Download_Files ($Url, $TargetFile) {
        # https://blogs.msdn.microsoft.com/jasonn/2008/06/13/downloading-files-from-the-internet-in-powershell-with-progress
        $uri = New-Object "System.Uri" "$Url"
        $request = [System.Net.HttpWebRequest]::Create($uri)
        $request.set_Timeout(15000) #15 second timeout
        $response = $request.GetResponse()
        # $totalLength = [System.Math]::Floor($response.get_ContentLength()/1024)
        $responseStream = $response.GetResponseStream()
        $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $TargetFile, Create
        $buffer = new-object byte[] 10KB
        $count = $responseStream.Read($buffer,0,$buffer.length)
        $downloadedBytes = $count
        while ($count -gt 0) {
            $targetStream.Write($buffer, 0, $count)
            $count = $responseStream.Read($buffer,0,$buffer.length)
            $downloadedBytes = $downloadedBytes + $count
            # Write-Progress -activity "Downloading file '$($Url.split('/') | Select-Object -Last 1)'" -status "Downloaded ($([System.Math]::Floor($downloadedBytes/1024))K of $($totalLength)K): " -PercentComplete ((([System.Math]::Floor($downloadedBytes/1024)) / $totalLength)  * 100)
        }
        # Write-Progress -activity "Finished downloading file '$($Url.split('/') | Select-Object -Last 1)'"
        $targetStream.Flush()
        $targetStream.Close()
        $targetStream.Dispose()
        $responseStream.Dispose()
    }
    # SCRIPT - Unzip
	function Unzip_File ($ZipFile,$DestFolder) {
	    Set-Alias sz "c:\program files\7-zip\7z.exe"
	    $command = "sz e -y $ZipFile `"-o$DestFolder`" "
	    $resultUnzip = Invoke-Expression $command | Out-Null
		if ($resultUnzip -contains "Everything is Ok") { $resultUnzip = $true } else { $resultUnzip = $false }
		return $resultUnzip
	}
    function Delete_Unwanted ($ListSkinsToDelete,$DataBase) {
        $LocalDatabase = "Local$DataBase"
        $dvLocalSkins = New-Object System.Data.DataView($dtsHT.$LocalDatabase)
        foreach ($skin in $ListSkinsToDelete) {
            $dvLocalSkins.RowFilter = "FileName = '$skin'"
            Remove-Item -Path $dvLocalSkins.FilePath -Force
        }
    }
     # LOGIC - Scan des skins locales
    function Scan_LocalSkins ($LocalPath,$Database,$Tags) {
        $LocalDatabase = "Local$Database"
        Get-ChildItem -Path "$LocalPath\*\*" -File | Where-Object {$Tags -contains $_.Name.Split('_')[1]} | ForEach-Object {
            $row = $dtsHT.$LocalDatabase.NewRow()
            $row["Tag"] = $_.Name.Split('_')[1]
            $row["FilePath"] = $_.FullName
            $row["FileName"] = $_.Name
            $row["FileDate"] = $_.LastWriteTime.ToString("yyyyMMdd")
            $row["PlaneType"] = $_.Name.Split('_')[0]
            $row["SkinCollection"] = $_.Name.Split('_')[0..3] -join ("_")
            $row["Quality"] = if ($_.Length -gt 6000000) { "4K" } else { "2K" }
            $dtsHT.$LocalDatabase.Rows.Add($row)
        }
    }
    # LOGIC - Scan des skins onlines
    function Scan_OnlineSkins ($baseUrl,$Tag,$Quality) {
        $URI = "$baseUrl/?dir=$Tag/$Quality"
        $HTML = Invoke-WebRequest -Uri $URI
        $fileNames = @()
        $fileSizes = @()
        ($HTML.ParsedHtml.getElementsByTagName('span') | Where-Object {$_.className -like 'file-name *'}).innerText | ForEach-Object { $fileNames += $_ }
        ($HTML.ParsedHtml.getElementsByTagName('span') | Where-Object {$_.className -like 'file-size *'}).innerText | ForEach-Object { $fileSizes += $_ }
        for ($i = 1; $i -lt $fileNames.count; $i++) {
            $fileName = $fileNames[$i].Trim()
            $fileSize = $fileSizes[$i].Trim()
            $row = $dtsHT.OnlineSkins.NewRow()
            $row["Tag"] = $Tag
            $row["FileUrl"] = ("{0}/{1}/{2}/{3}" -f $baseUrl,$Tag,$Quality,$fileName)
            $row["FileName"] = $fileName
            $row["FileSize"] = $fileSize
            $row["FileDate"] = $fileName.Split('_')[4].Substring(0,$fileName.Split('_')[4].Length-3)
            $row["PlaneType"] = $fileName.Split('_')[0]
            $row["SkinCollection"] = ($fileName.Split('_')[0..3] -join ("_"))
            $row["Quality"] = $Quality
            $dtsHT.OnlineSkins.Rows.Add($row)
        }
    }
    # LOGIC - Comparaison des SkinCollections Locales / Onlines
    function Compare_Collections ($QualityPref,$Database) {
        $LocalDatabase = "Local$Database"
        # Uniques SkinCollection & Quality
        $dvLocalSkins = New-Object System.Data.DataView($dtsHT.$LocalDatabase)
        $uniquesDt = $dvLocalSkins.ToTable($true, "SkinCollection","Quality")
        # Nouvelle DT avec FileDate
        $dtLocalCollection = New-Object System.Data.Datatable
        [Void]$dtLocalCollection.Columns.Add("SkinCollection")
        [Void]$dtLocalCollection.Columns.Add("Quality")
        [Void]$dtLocalCollection.Columns.Add("FileDate")
        for ($i = 0; $i -lt $uniquesDt.FileName.Length; $i++) {
            $findSkin = $uniquesDt.SkinCollection[$i]
            $dvLocalSkins.RowFilter = "SkinCollection = '$findSkin'"
            $date = ($dvLocalSkins | Sort-Object FileDate -Descending | Select-Object FileDate -First 1 ).FileDate
            $row = $dtLocalCollection.NewRow()
            $row["SkinCollection"] = $findSkin
            $row["Quality"] = $uniquesDt.Quality[$i]
            $row["FileDate"] = $date
            $dtLocalCollection.Rows.Add($row)
        }
        $dvOnlineSkins = New-Object System.Data.DataView($dtsHT.OnlineSkins)
        $listToDownload = @()
        # Comparaison des skins locales avec les skins Onlines
        foreach ($row in $dtLocalCollection) {
            $findSkin = $row.SkinCollection
            $fileDate = $row.FileDate
            $fileQuality = $row.Quality
            $dvOnlineSkins.RowFilter = "SkinCollection = '$findSkin'"
            # Si on a les deux qualités
            if ($dvOnlineSkins.Count -eq 2) {
                foreach ($Online in $dvOnlineSkins) {
                    # Si la qualité matche les préférences
                    if ($Online.Quality -eq $QualityPref) {
                        # Si la skin locale matche les préférences
                        if ($fileQuality -eq $QualityPref) {
                            # On la récupère si elle est plus récente
                            if ($Online.FileDate -gt $fileDate) {
                                $listToDownload += $Online.FileUrl
                            }
                        }
                        # sinon on télécharge celle qui correspond aux préférences
                        else { $listToDownload += $Online.FileUrl }
                    }
                }
            }
            # On a qu'une seule qualité
            else {
                # Si elle correspond aux préférences
                if ($dvOnlineSkins.Quality -eq $QualityPref) {
                    # Si la qualité correspond à celle qu'on a déjà en local
                    if ($dvOnlineSkins.Quality -eq $fileQuality) {
                        # Si elle est plus récente on la prend
                        if ($dvOnlineSkins.FileDate -gt $fileDate) {
                            $listToDownload += $dvOnlineSkins.FileUrl
                        }
                    }
                    # Sinon ben on la prend
                    else { $listToDownload += $dvOnlineSkins.FileUrl }
                }
                # Si elle correspond PAS aux préférences
                else {
                    # Si la qualité de la locale ne correspond non plus aux préférences
                    if ($fileQuality -ne $QualityPref) {
                        # On la prend si plus récente
                        if ($dvOnlineSkins.FileDate -gt $fileDate) { $listToDownload += $dvOnlineSkins.FileUrl }
                    }
                }
            }
        }
        # # Comparaison des skins Onlines avec les skins Locales pour trouver celles qu'on ne possède pas
        foreach ($row in $dtsHT.OnlineSkins) {
            # Si on ne possède pas la collection en local
            if ($dtsHT.$LocalDatabase.SkinCollection -notcontains $row.SkinCollection) {
                $findSkin = $row.SkinCollection
                $dvOnlineSkins.RowFilter = "SkinCollection = '$findSkin'"
                # Si on a les 2 qualités
                if ($dvOnlineSkins.Count -eq 2) {
                    foreach ($Online in $dvOnlineSkins) {
                        # Si la qualité correspond aux préférences on la prend
                        if ($Online.Quality -eq $QualityPref) { $listToDownload += $Online.FileUrl }
                    }
                }
                else { $listToDownload += $dvOnlineSkins.FileUrl }
            }
        }
        $listToDownload = $listToDownload | Sort-Object -Unique
        # Confection de la DataBase d'Update
        $DatabaseToUpdate = "SkinsToUpdate$Database"
        foreach ($file in $listToDownload) {
            $dvOnlineSkins.RowFilter = "FileUrl = '$file'"
            $row = $dtsHT.$DatabaseToUpdate.NewRow()
            $row["Tag"] = $dvOnlineSkins.Tag
            $row["FileUrl"] = $dvOnlineSkins.FileUrl
            $row["FileName"] = $dvOnlineSkins.FileName
            $row["FileSize"] = $dvOnlineSkins.FileSize
            $row["FileDate"] = $dvOnlineSkins.FileDate
            $row["PlaneType"] = $dvOnlineSkins.PlaneType
            $row["SkinCollection"] = $dvOnlineSkins.SkinCollection
            $row["Quality"] = $dvOnlineSkins.Quality
            $dtsHT.$DatabaseToUpdate.Rows.Add($row)
        }
    }
    # LOGIC - Vérification des skins à delete
    function Check_ToDelete ($baseUrl,$Database,$Tag) {
        $LocalDatabase = "Local$Database"
        $ToDelete = @()
        $URI = "$baseUrl/$Tag/todelete.txt"
        $SkinsToDelete = (Invoke-WebRequest -Uri $URI).Content
        if ($SkinsToDelete) {
            $dvLocalSkins = New-Object System.Data.DataView($dtsHT.$LocalDatabase)
            $SkinsToDelete.Replace("`n",";").Split(";") | ForEach-Object {
                $dvLocalSkins.RowFilter = "FileName = '$_'"
                if ($dvLocalSkins.Count -gt 0) {
                    $ToDelete += $_
                }
            }
            return $ToDelete
        }
    }

#endregion

#region Initialisations

    # Base Script Config
    $version        = 1.4
    $baseUrl        = "https://www.lesirreductibles.com/irreskins"
    $fileConf       = "config.json"

    # WPF Assemblies
	Add-Type -AssemblyName PresentationFramework            
	Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    # TLS 1.2
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Import XAML
    $gui 			= [hashtable]::Synchronized(@{})
	$reader     	= (New-Object Xml.XmlNodeReader $xaml)
    $gui.Window		= [Windows.Markup.XamlReader]::Load($reader)
    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object { $gui.Add($_.Name,$gui.Window.FindName($_.Name)) }
    # Init ScriptHT
    $ScriptHT                       = [hashtable]::Synchronized(@{})
    $ScriptHT.Config                = @{}
    $ScriptHT.baseUrl               = $baseUrl
    $ScriptHT.ScriptPath            = $PSScriptRoot
    $ScriptHT.ConfigFile            = $env:APPDATA + "\IRREskins\config.json"
    $ScriptHT.LogsFile              = $env:APPDATA + "\IRREskins\irreskins.log"
    $ScriptHT.DwlFolder             = $env:APPDATA + "\IRREskins\Downloads"
    $ScriptHT.Installations         = @("StandAlone","Steam")
    $dtsHT                          = [hashtable]::Synchronized(@{})
    #bases64
	$base64img_icon = New-Object System.Windows.Media.Imaging.BitmapImage
	$base64img_icon.BeginInit()
	$base64img_icon.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String("iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAABmJLR0QA/wD/AP+gvaeTAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAB3RJTUUH4QwLECEGeucbIQAABUtJREFUaN7VWk9IG10Q/7Vd06ityWEJEkRFJGjIQSPYsIQoEsX2UPSiICI9eRYRycGjSE/Sg4ZSREqp0ENR0SAhKCiI9g9UsSJYStAgbQwSmpLquq5OD9/nw/V/kl1TBx68fW/27cy8mXkz8/YOERE0hL29PWRmZmq2/l2tFt7f30dpaSnsdjtisZh2EiKNwOFwEAACQBaLhcLhsCbfgdbEHze3202yLN8OBgKBAOl0OgUDb9++vT07QEQ0PDzMiH/+/LlWnyHVjHhpaUnxzPM865tMJsXc5OQkDg4O/h0vVFFRAbvdjo8fP547f3R0xPqDg4N4+vQpWlpacHh4mF4vJEkS1dfXK3Td5XKR0+kkq9XKxoqLi8npdJLT6VTgPnv2LGXDTomB7u7uM94m0dbf358+BqLRqELSAEgQBHI4HFRSUsLGioqKSBAEEgRBgVtXV0eiKKbXC8XjceJ5ngCQ3+8nWZZJFEV6//49I/Tly5dERCSKIvX19alGvGpuVJIkCgQCirGJiQnGwPDwsGLO6/WqQjwREZeqE/jz5w+CwSDMZjO+fv3Kxjc2Nlg/FAphfX0dkiQBAFwuF0KhEPLy8lIP9JLhemNjg7q7u0mv16dsxCaTifr7+2lnZ+dmVGh8fJyMRmPKhJ9uhYWFSQV8dxLJBzY3N2G1WrG7uwsAyMvLg9VqRVFRkeKwOlYbv98PAKiurobVaoUsywqcb9++YW1tDZFIBACg0+kQjUaRnZ2tjQrZbDYmsc7Ozku3fWpqiuG+fv36QrxwOExNTU0Mt6mpSRsV+vDhA/uIzWa7Ev+kFxoaGroS32KxMJtIxB6uHQv5fD7WHxgYUD2x8ng8AIBIJIKdnR31g7lwOMz6VVVVqjOQn58PjuMYE6ozoNfrWf/nz5+qMyCKIjPynJwc9RlwOp2s39vbqzoD4+PjTFC5ubnaeKGsrCwCQBzHnQkdUjHiN2/eKIJBSZK0OchOukYA1NHRQZFIJGkGgsEgtba2KtacnZ3V9iT2eDyqn8LH7dWrVzcTjfr9fqZOajSe5+nTp09JxUIJhRIn4eDgAD6fD8vLy5BlGaurq5iYmGDzZrMZP378AADYbDasrq6yuZqaGrhcLnAch8rKStTW1qa3Mre7u0tms5lJ1G630+joqEI1urq62DPHcQnrumYJTSgUYhkZAKqurr7QiFtaWtjYgwcPaGVlJT0JzeLiIqLRKCRJQmdnJzv6y8rKEAgELnxvZGQERqMRXq8X8XgcDocDQ0NDyMnJAc/zePTo0c2o0HmGKAjCtd1oQ0PDmfeNRiP9+vVL+8rceRU1u92O6enpa68xNjaG9vZ2xVhWVlb6KnM9PT0J57YOhyM9pcWMjIxLA73rwmmJi6KIu3cTl2dSRizLMuLx+H8LcFxiKeD/0NzcjCdPnrBUVK/X4/79+0xNzxNU0jswMzODxsZG9nzv3j0YDAYYDIakiD+Ghw8fsnWOid/e3kZdXR0+f/6szg7Mzc3B7XYDAMrLyy91kyfh9+/frB+PxxGLxVhd6CLQ6XQQBAHBYBD19fX48uULCgoKknej8/PzZ25abrKZTKYrD7tLVWhra+tKqWkJkUgEW1tbyatQc3Mzvn//jp6eHgBAcXExPB7PmfrOebCysgKv1wsAaGtrg8vluvI9nU6H3t5eBINBcByHd+/e4fHjx6mfxC9evCCTyZTQCZloWeUY1tbWiOO4a18KanbJd1l1+p+85DsNJ2s7iZRJ/omb+lt/T1xWVna7b+pFUVTcnWn5r0TSOfFVEIvFYLFYwPM8FhYWYDAYNDGBv+RUPu9Ne3B4AAAAAElFTkSuQmCC")
	$base64img_icon.EndInit()
	$base64img_icon.Freeze()
    # Local Download folders
    if (!(Test-Path "$($ScriptHT.DwlFolder)\2K")) { New-Item -Path "$($ScriptHT.DwlFolder)\2K" -ItemType Directory -Force }
    if (!(Test-Path "$($ScriptHT.DwlFolder)\4K")) { New-Item -Path "$($ScriptHT.DwlFolder)\4K" -ItemType Directory -Force }
    # Si une configuration locale existe déjà on la récupère en interne
    if (Test-Path $ScriptHT.ConfigFile) {
        $LocalConfig = (Get-Content $ScriptHT.ConfigFile) -join "`n" | ConvertFrom-Json
        $LocalConfig | ForEach-Object {
            $ScriptHT.Config.StandAlonePath		= $LocalConfig.StandAlonePath
            $ScriptHT.Config.SteamPath			= $LocalConfig.SteamPath
            $ScriptHT.Config.SZip				= $LocalConfig.SZip
            $ScriptHT.Config.Pref_StandAlone	= $LocalConfig.Pref_StandAlone
            $ScriptHT.Config.Pref_Steam			= $LocalConfig.Pref_Steam
            $ScriptHT.Config.Pref_Quality		= $LocalConfig.Pref_Quality
            $ScriptHT.Config.Pref_KeepFiles		= $LocalConfig.Pref_KeepFiles
        }
        # On Check les Paths
        Check_Path_StandAlone -Path $LocalConfig.StandAlonePath
        Check_Path_Steam -Path $LocalConfig.SteamPath
        Check_Path_SZip -Path $LocalConfig.SZip
    }
    # Sinon on en crée une par défault
    else {
        Check_Path_StandAlone
        Check_Path_Steam
        Check_Path_SZip
        # $ScriptHT.Config.Pref_StandAlone & $ScriptHT.Config.Pref_Steam sont déterminés par les Check_Path
        $ScriptHT.Config.Pref_Quality		= "4K"
        $ScriptHT.Config.Pref_KeepFiles		= $true
    }
    # On applique la config et on la sauvegarde
    Apply_Config

#endregion

#region Main

    # Action à l'initialisation de la fenêtre
    $gui.Window.Add_SourceInitialized({
        # Icon
        $gui.Window.Icon = $base64img_icon
        # Version
        $gui.LB_Version.Content = [string]"v$version"
        # Init Boutons
        $gui.BT_Apply.IsEnabled = $false
        # Vérifications
        ## Vérification de la version Powershell
        $psverMajor = $PsVersionTable.PSVersion.Major
        $psverMinor = $PsVersionTable.PSVersion.Minor
        New_Line -Text "Version de Powershell : $psverMajor.$psverMinor" -Paragraph "P_Execution" -Foreground "Blue"
        if ($psverMajor -lt 3) {
            New_Line -Text "IRREskins n'est pas compatible avec votre version de Powershell, veuillez installer la version 3 ou supérieure" -Paragraph "P_Execution" -Foreground "Red"
        }
        else {
            ## Vérification de la connexion Online
            Try {
                $OnlineConfig = Invoke-WebRequest -Uri "$baseUrl/$fileConf" | ConvertFrom-Json
                New_Line -Text "Récupération de la config Online OK" -Paragraph "P_Execution" -Foreground "Green"
                ## Add Tags
                $ScriptHT.Tags = $OnlineConfig.Tags
                ## Check Version
                if ($OnlineConfig.version -gt $version) { Maj_Script -Version $OnlineConfig.version -InstallFile $OnlineConfig.installFile }
                else { New_Line -Text "Pas de nouvelle version d'IRREskins`r`nCliquez sur Scanner pour vérifier les mises à jour" -Paragraph "P_Execution" }
            }
            Catch {
                New_Line -Text "Echec de la récupération de la config Online :(" -Paragraph "P_Execution" -Foreground "Red"
                New_Line -Break -Paragraph "P_Execution"
                New_Line -Text "Réessayez de lancer IRREskins après avoir lancé Internet Explorer." -Paragraph "P_Execution"
                New_Line -Break -Paragraph "P_Execution"
                New_Line -Text "Si ça ne fonctionne toujours pas, désinstallez IRREskins et réinstallez le en récupérant le dernier Setup disponible en ligne." -Paragraph "P_Execution"
                $gui.BT_Scan.IsEnabled = $false
            }
        }
    })
    # Action sur le bouton de recherche du path de BOS StandAlone
    $gui.BT_StandAlone.add_Click({
        $SelectedFolder = Select_Folder -Message "Sélectionnez le dossier racine de BOS (non Steam)"
        $TestFolder = "$SelectedFolder\bin\game\il-2.exe"
        if (Test-Path $TestFolder) { Check_Path_StandAlone -Path $TestFolder }
        else { Check_Path_StandAlone -Path $null }
        Apply_Config
    })
    # Action sur le bouton de recherche du path de Steam
    $gui.BT_Steam.add_Click({
        $SelectedFolder = Select_Folder -Message "Sélectionnez le dossier racine de Steam ou le dossier racine de BOS (version Steam) si vous ne l'avez pas installé sous SteamApps"
        # Test si BOS est bien dans SteamApps & Hors de SteamApps
        $TestFolderINSteamApps = "$SelectedFolder\steamApps\common\il-2 sturmovik battle of stalingrad\bin\game\il-2.exe"
        $TestFolderOUTSteamApps = "$SelectedFolder\bin\game\il-2.exe"
        if (Test-Path $TestFolderINSteamApps) { Check_Path_Steam -Path $TestFolderINSteamApps }
        elseif (Test-Path $TestFolderOUTSteamApps) { Check_Path_Steam -Path $TestFolderOUTSteamApps }
        else { Check_Path_Steam -Path $null }
        Apply_Config
    })
    # Action sur le bouton de recherche du path de 7-Zip
    $gui.BT_SZip.add_Click({
        $SelectedFolder = Select_Folder -Message "Sélectionnez le dossier racine de 7-Zip"
        $TestFolder = "$SelectedFolder\7z.exe"
        if (Test-Path $TestFolder) { Check_Path_SZip -Path $SelectedFolder }
        else { Check_Path_SZip -Path $null }
        Apply_Config
    })
    # Action sur le bouton d'ouverture du lien 7-Zip
    $gui.BT_LinkSZip.add_Click({
        $url = "http://www.7-zip.org"
        (New-Object -Com Shell.Application).Open($url)
    })
    # Action sur le bouton d'ouverture du répertoire de Download
    $gui.BT_OpenFilesFolder.add_Click({
        Invoke-Item $ScriptHT.DwlFolder
    })
    # Check StandAlone 
    $gui.CB_StandAlone.add_Click({
        # Si Steam Config OK et Steam CB Non coché Alors StandAlone CB forcé OK 
        if (($ScriptHT.Config.SteamPath -ne $null) -and ($gui.CB_Steam.IsChecked -eq $false)) { $gui.CB_StandAlone.IsChecked = $true }
        else {
            if ($gui.CB_StandAlone.IsChecked) { $ScriptHT.Config.Pref_StandAlone = $true }
            else { $ScriptHT.Config.Pref_StandAlone = $false }
            Apply_Config
        }
    })
    # Check Steam
    $gui.CB_Steam.add_Click({
        # Si StandAlone Config OK et StandAlone CB Non coché Alors Steam CB forcé OK
        if (($ScriptHT.Config.StandAlonePath -ne $null) -and ($gui.CB_StandAlone.IsChecked -eq $false)) { $gui.CB_Steam.IsChecked = $true }
        else {
            if ($gui.CB_Steam.IsChecked) { $ScriptHT.Config.Pref_Steam = $true }
            else { $ScriptHT.Config.Pref_Steam = $false }
            Apply_Config
        }
    })
    # Check KeepFiles
    $gui.CB_KeepFiles.add_Click({
        if ($gui.CB_KeepFiles.IsChecked) { $ScriptHT.Config.Pref_KeepFiles = $true }
        else { $ScriptHT.Config.Pref_KeepFiles = $false }
        Apply_Config
    })
    # Slider
    $gui.SL_Quality.add_ValueChanged({
        switch ($gui.SL_Quality.Value) {
            0 { $ScriptHT.Config.Pref_Quality = "2K"; break }
            1 { $ScriptHT.Config.Pref_Quality = "4K"; break }
        }
        Apply_Config
    })
    # Action sur le bouton de Scan
    $gui.BT_Scan.add_Click({
        # DataTables HT
        ## OnlineSkins
        $dtsHT.OnlineSkins = New-Object System.Data.Datatable
        [Void]$dtsHT.OnlineSkins.Columns.Add("Tag")
        [Void]$dtsHT.OnlineSkins.Columns.Add("FileUrl")
        [Void]$dtsHT.OnlineSkins.Columns.Add("FileName")
        [Void]$dtsHT.OnlineSkins.Columns.Add("FileSize")
        [Void]$dtsHT.OnlineSkins.Columns.Add("FileDate")
        [Void]$dtsHT.OnlineSkins.Columns.Add("PlaneType")
        [Void]$dtsHT.OnlineSkins.Columns.Add("SkinCollection")
        [Void]$dtsHT.OnlineSkins.Columns.Add("Quality")
        ## LocalStandAlone
        $dtsHT.LocalStandAlone = New-Object System.Data.Datatable
        [Void]$dtsHT.LocalStandAlone.Columns.Add("Tag")
        [Void]$dtsHT.LocalStandAlone.Columns.Add("FilePath")
        [Void]$dtsHT.LocalStandAlone.Columns.Add("FileName")
        [Void]$dtsHT.LocalStandAlone.Columns.Add("FileDate")
        [Void]$dtsHT.LocalStandAlone.Columns.Add("PlaneType")
        [Void]$dtsHT.LocalStandAlone.Columns.Add("SkinCollection")
        [Void]$dtsHT.LocalStandAlone.Columns.Add("Quality")
        ## LocalSteam
        $dtsHT.LocalSteam = New-Object System.Data.Datatable
        [Void]$dtsHT.LocalSteam.Columns.Add("Tag")
        [Void]$dtsHT.LocalSteam.Columns.Add("FilePath")
        [Void]$dtsHT.LocalSteam.Columns.Add("FileName")
        [Void]$dtsHT.LocalSteam.Columns.Add("FileDate")
        [Void]$dtsHT.LocalSteam.Columns.Add("PlaneType")
        [Void]$dtsHT.LocalSteam.Columns.Add("SkinCollection")
        [Void]$dtsHT.LocalSteam.Columns.Add("Quality")
        ## SkinsToUpdateStandAlone
        $dtsHT.SkinsToUpdateStandAlone = New-Object System.Data.Datatable
        [Void]$dtsHT.SkinsToUpdateStandAlone.Columns.Add("Tag")
        [Void]$dtsHT.SkinsToUpdateStandAlone.Columns.Add("FileUrl")
        [Void]$dtsHT.SkinsToUpdateStandAlone.Columns.Add("FileName")
        [Void]$dtsHT.SkinsToUpdateStandAlone.Columns.Add("FileSize")
        [Void]$dtsHT.SkinsToUpdateStandAlone.Columns.Add("FileDate")
        [Void]$dtsHT.SkinsToUpdateStandAlone.Columns.Add("PlaneType")
        [Void]$dtsHT.SkinsToUpdateStandAlone.Columns.Add("SkinCollection")
        [Void]$dtsHT.SkinsToUpdateStandAlone.Columns.Add("Quality")
        ## SkinsToUpdateSteam
        $dtsHT.SkinsToUpdateSteam = New-Object System.Data.Datatable
        [Void]$dtsHT.SkinsToUpdateSteam.Columns.Add("Tag")
        [Void]$dtsHT.SkinsToUpdateSteam.Columns.Add("FileUrl")
        [Void]$dtsHT.SkinsToUpdateSteam.Columns.Add("FileName")
        [Void]$dtsHT.SkinsToUpdateSteam.Columns.Add("FileSize")
        [Void]$dtsHT.SkinsToUpdateSteam.Columns.Add("FileDate")
        [Void]$dtsHT.SkinsToUpdateSteam.Columns.Add("PlaneType")
        [Void]$dtsHT.SkinsToUpdateSteam.Columns.Add("SkinCollection")
        [Void]$dtsHT.SkinsToUpdateSteam.Columns.Add("Quality")
        # Ajoute les fonctions au Runspace
        $FunctionList = @("Scan_OnlineSkins","Scan_LocalSkins","Check_ToDelete","Compare_Collections","Display_ScanResults","New_Line")
        $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($Function in $FunctionList) {
            $FunctionDefinition = Get-Content Function:\$Function
            $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $Function, $FunctionDefinition
            $initialSessionState.Commands.Add($SessionStateFunction)
        }
        # Création du Runspace
        $Runspace = [runspacefactory]::CreateRunspace($initialSessionState)
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions	 = "ReuseThread"
        $Runspace.Open()
        $Runspace.SessionStateProxy.SetVariable("gui",$gui)
        $Runspace.SessionStateProxy.SetVariable("ScriptHT",$ScriptHT)
        $Runspace.SessionStateProxy.SetVariable("dtsHT",$dtsHT)
        # Code d'exécution du Runspace
        $Code = {
            # Récupération des skins locales
            $gui.Window.Dispatcher.invoke("Normal",[action]{
                $gui.BT_Scan.IsEnabled = $false
                $gui.BT_Apply.IsEnabled = $false
                $gui.GP_Paths.IsEnabled = $false
                $gui.GP_Preferences.IsEnabled = $false
                $gui.P_Execution.Inlines.Clear()
                $gui.PB_Progress.Value = 0
                New_Line -Text "Scan des skins locales ..." -Paragraph "P_Execution"
            })
            if ($ScriptHT.Config.Pref_StandAlone) { Scan_LocalSkins -localPath $ScriptHT.Config.StandAlonePath -Database "StandAlone" -Tags $ScriptHT.Tags }
            if ($ScriptHT.Config.Pref_Steam) { Scan_LocalSkins -localPath $ScriptHT.Config.SteamPath -Database "Steam" -Tags $ScriptHT.Tags }
            # Récupération de la liste des skins Online
            $gui.Window.Dispatcher.invoke("Normal",[action]{ New_Line -Text "Scan des skins onlines ..." -Paragraph "P_Execution" })
            foreach ($Tag in $ScriptHT.Tags) {
                Scan_OnlineSkins $ScriptHT.baseUrl $Tag "2K"
                Scan_OnlineSkins $ScriptHT.baseUrl $Tag "4K"
            }
            # Récupération des skins à delete
            $gui.Window.Dispatcher.invoke("Normal",[action]{ New_Line -Text "Récupération de la liste des skins obsolètes ..." -Paragraph "P_Execution" })
            $ScriptHT.ToDeleteStandAlone = @()
            $ScriptHT.ToDeleteSteam = @()
            foreach ($Tag in $ScriptHT.Tags) {
                if ($ScriptHT.Config.Pref_StandAlone) { $ScriptHT.ToDeleteStandAlone += Check_ToDelete -baseUrl $ScriptHT.baseUrl -Database "StandAlone" -Tag $Tag }
                if ($ScriptHT.Config.Pref_Steam) { $ScriptHT.ToDeleteSteam += Check_ToDelete -baseUrl $ScriptHT.baseUrl -Database "Steam" -Tag $Tag }
            }
            # Comparaisons, skins à mettre à jour
            $gui.Window.Dispatcher.invoke("Normal",[action]{ New_Line -Text "Analyse des données ..." -Paragraph "P_Execution" })
            if ($ScriptHT.Config.Pref_StandAlone) { Compare_Collections -QualityPref $ScriptHT.Config.Pref_Quality -Database "StandAlone" }
            if ($ScriptHT.Config.Pref_Steam) { Compare_Collections -QualityPref $ScriptHT.Config.Pref_Quality -Database "Steam" }
            # Reset RTB
            $gui.Window.Dispatcher.invoke("Normal",[action]{ $gui.P_Execution.Inlines.Clear() })
            # Résultats StandAlone
            if ($ScriptHT.Config.Pref_StandAlone) {
                $gui.Window.Dispatcher.invoke("Normal",[action]{
                    Display_ScanResults -Database "StandAlone"
                })
                $StandAloneDisplay = $true
            }
            # Résultats Steam
            if ($ScriptHT.Config.Pref_Steam) {
                $gui.Window.Dispatcher.invoke("Normal",[action]{
                    if ($StandAloneDisplay) { New_Line -Break -Paragraph "P_Execution" }
                    Display_ScanResults -Database "Steam"
                })
            }
            # Activation du Bouton d'installation
            $gui.Window.Dispatcher.invoke("Normal",[action]{
                $gui.BT_Scan.IsEnabled = $true
                $gui.GP_Paths.IsEnabled = $true
                $gui.GP_Preferences.IsEnabled = $true
                if (($ScriptHT.ToDeleteStandAlone -gt 0) -or ($ScriptHT.ToDeleteSteam -gt 0) -or ($dtsHT.SkinsToUpdateStandAlone.Item.Count -gt 0) -or ($dtsHT.SkinsToUpdateSteam.Item.Count -gt 0)) {
                        $gui.BT_Apply.IsEnabled = $true
                }
            })
        }
        # Invocation du Runspace
        $PSinstance	 = [powershell]::Create().AddScript($Code)
        $PSinstance.Runspace = $Runspace
        $PSinstance.BeginInvoke()
    })
    # Action sur le bouton Apply
    $gui.BT_Apply.add_Click({
        # Ajoute les fonctions au Runspace
        $FunctionList = @("Delete_Unwanted","Download_Files","Unzip_File","New_Line")
        $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($Function in $FunctionList) {
            $FunctionDefinition = Get-Content Function:\$Function
            $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $Function, $FunctionDefinition
            $initialSessionState.Commands.Add($SessionStateFunction)
        }
        # Création du Runspace
        $Runspace = [runspacefactory]::CreateRunspace($initialSessionState)
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions	 = "ReuseThread"
        $Runspace.Open()
        $Runspace.SessionStateProxy.SetVariable("gui",$gui)
        $Runspace.SessionStateProxy.SetVariable("ScriptHT",$ScriptHT)
        $Runspace.SessionStateProxy.SetVariable("dtsHT",$dtsHT)
        # Code d'exécution du Runspace
        $Code = {
            # Reset RTB & PB & GP & désactivation des bouton Apply,Scanner & Quit
            $gui.Window.Dispatcher.invoke("Normal",[action]{
                $gui.PB_Progress.Value = 0
                $gui.GP_Paths.IsEnabled = $false
                $gui.GP_Preferences.IsEnabled = $false
                $gui.P_Execution.Inlines.Clear()
                $gui.BT_Scan.IsEnabled = $false
                $gui.BT_Apply.IsEnabled = $false
                $gui.BT_Quit.IsEnabled = $false
            })
            # Suppression des skins obsolètes
            foreach ($Installation in $ScriptHT.Installations) {
                $list = "ToDelete$Installation"
                if ($ScriptHT.$list.Count -gt 0) {
                    $Database = "StandAlone"
                    $gui.Window.Dispatcher.invoke("Normal",[action]{
                        New_Line -Text "Suppression des skins obsolètes de l'installation $Installation ..." -Bold -Paragraph "P_Execution"
                    })
                    Delete_Unwanted -ListSkinsToDelete $ScriptHT.$list -DataBase $Installation
                    $SkinsDeleted = $true
                }
            }
            # Téléchargement
            $ProgressMaxValue = $dtsHT.SkinsToUpdateStandAlone.Item.Count + $dtsHT.SkinsToUpdateSteam.Item.Count
            $ProgressCurrent = 0
            if (($dtsHT.SkinsToUpdateStandAlone.Item.Count -gt 0) -or ($dtsHT.SkinsToUpdateSteam.Item.Count -gt 0)) {
                $gui.Window.Dispatcher.invoke("Normal",[action]{
                    if ($SkinsDeleted) { New_Line -Break -Paragraph "P_Execution" }
                    New_Line -Text "Téléchargement des Collections ..." -Bold -Paragraph "P_Execution"
                    $gui.PB_Progress.Maximum = $ProgressMaxValue
                    $gui.PB_Progress.Value = 0
                })
                foreach ($Installation in $ScriptHT.Installations) {
                    $Database = "SkinsToUpdate$Installation"
                    foreach ($row in $dtsHT.$Database) {
                        $destFile = $ScriptHT.DwlFolder + "\" + $row.Quality + "\" + $row.FileName
                        $ProgressCurrent++
                        if (Test-Path $destFile) {
                            $gui.Window.Dispatcher.invoke("Normal",[action]{
                                New_Line -Text "Fichier déjà téléchargé : $($row.FileName)" -Paragraph "P_Execution"
                                $gui.PB_Progress.Value = $ProgressCurrent
                            })
                        }
                        else {
                            $gui.Window.Dispatcher.invoke("Normal",[action]{
                                New_Line -Text "Téléchargement en cours : $($row.FileName) ($($row.FileSize)) ..." -Paragraph "P_Execution"
                            })
                            Download_Files -Url $row.FileUrl -TargetFile $destFile
                            $gui.Window.Dispatcher.invoke("Normal",[action]{
                                $gui.PB_Progress.Value = $ProgressCurrent
                            })
                        }   
                    }
                }
            }
            # Unzip
            $ProgressMaxValue = $dtsHT.SkinsToUpdateStandAlone.Item.Count + $dtsHT.SkinsToUpdateSteam.Item.Count
            $ProgressCurrent = 0
            if (($dtsHT.SkinsToUpdateStandAlone.Item.Count -gt 0) -or ($dtsHT.SkinsToUpdateSteam.Item.Count -gt 0)) {
                $gui.Window.Dispatcher.invoke("Normal",[action]{
                    New_Line -Break -Paragraph "P_Execution"
                    New_Line -Text "Installation des Collections ..." -Bold -Paragraph "P_Execution"
                    $gui.PB_Progress.Maximum = $ProgressMaxValue
                    $gui.PB_Progress.Value = 0
                })
                foreach ($Installation in $ScriptHT.Installations) {
                    $Database = "SkinsToUpdate$Installation"
                    foreach ($row in $dtsHT.$Database) {
                        $ZipFile = ("{0}\{1}\{2}" -f $ScriptHT.DwlFolder,$row.Quality,$row.FileName)
                        $InstallationPath = $Installation + "Path"
                        $DestFolder = $ScriptHT.Config.$InstallationPath + "\" + $row.PlaneType
                        $ProgressCurrent++
                        $gui.Window.Dispatcher.invoke("Normal",[action]{
                            New_Line -Text "Installation $Installation : $($row.FileName) ..." -Paragraph "P_Execution"
                        })
                        $testUnzip = Unzip_File -ZipFile $ZipFile -DestFolder $DestFolder
                        $gui.Window.Dispatcher.invoke("Normal",[action]{
                            if ($testUnzip -eq $false) {
                                New_Line -Text "Erreur lors de l'installation de la collection" -Paragraph "P_Execution" -Foreground "Red"
                            }
                            $gui.PB_Progress.Value = $ProgressCurrent
                        })
                    }
                }
            }
            # Suppression des 7-Zip téléchargés
            if ($ScriptHT.Config.Pref_KeepFiles -eq $false) {
                $gui.Window.Dispatcher.invoke("Normal",[action]{
                    New_Line -Break -Paragraph "P_Execution"
                    New_Line -Text "Suppression des fichiers téléchargés ..." -Bold -Paragraph "P_Execution"
                })
                foreach ($Installation in $ScriptHT.Installations) {
                    $Database = "SkinsToUpdate$Installation"
                    foreach ($row in $dtsHT.$Database) {
                        $FileToDelete = $ScriptHT.DwlFolder + "\" + $row.Quality + "\" + $row.FileName
                        Remove-Item -Path $FileToDelete -Force
                    }
                }
            }
            # Réactivation du bouton Scan
            $gui.Window.Dispatcher.invoke("Normal",[action]{
                New_Line -Break -Paragraph "P_Execution"
                New_Line -Text "Opération terminées !" -Bold -Foreground "Green" -Paragraph "P_Execution"
                $gui.GP_Paths.IsEnabled = $true
                $gui.GP_Preferences.IsEnabled = $true
                $gui.BT_Scan.IsEnabled = $true
                $gui.BT_Quit.IsEnabled = $true
            })
        }
        # Invocation du Runspace
        $PSinstance	 = [powershell]::Create().AddScript($Code)
        $PSinstance.Runspace = $Runspace
        $PSinstance.BeginInvoke()
    })
    # Action sur le bouton Quitter
    $gui.BT_Quit.add_Click({
        $gui.window.Close()
    })
    ## Garbage Collector
    $gui.Window.Add_Closed({
        Save_Config
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
    })

#endregion

#region Launch GUI

    $gui.Window.ShowDialog() | Out-Null

#endregion
