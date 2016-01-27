$checkbox = $null
Function CheckBox{
#==============================================================================================
# XAML Code - Imported from Visual Studio Express WPF Application
#==============================================================================================
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MainWindow" Height="395" Width="463">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="130"/>
            <RowDefinition Height="338"/>
        </Grid.RowDefinitions>

        <TextBlock Name="MaterialText" HorizontalAlignment="Left" Margin="10,45,0,0" TextWrapping="Wrap" Text="Material" VerticalAlignment="Top" Height="22" Width="63" FontSize="16"/>
        <TextBlock Name="CAMText" HorizontalAlignment="Left" Margin="70,88,0,0" TextWrapping="Wrap" Text="CAM" VerticalAlignment="Top" FontSize="16"/>
        <TextBlock Name="Sawing" HorizontalAlignment="Left" Margin="70,114,0,0" TextWrapping="Wrap" Text="Sawing" VerticalAlignment="Top" FontSize="16" Height="23" Grid.RowSpan="2"/>
        <TextBlock Name="StandardTurning" HorizontalAlignment="Left" Margin="70,10,0,0" TextWrapping="Wrap" Text="Standard Turning" VerticalAlignment="Top" Grid.Row="1" FontSize="16"/>
        <TextBlock Name="CNCTurning" HorizontalAlignment="Left" Margin="70,36,0,0" TextWrapping="Wrap" Text="CNC Turning" VerticalAlignment="Top" Grid.Row="1" FontSize="16"/>
        <TextBlock Name="StandardMilling" HorizontalAlignment="Left" Margin="70,62,0,0" TextWrapping="Wrap" Text="Standard Milling" VerticalAlignment="Top" Grid.Row="1" FontSize="16"/>
        <TextBlock Name="MultiaxialMilling" HorizontalAlignment="Left" Margin="70,88,0,0" TextWrapping="Wrap" Text="Multiaxial Milling" VerticalAlignment="Top" Grid.Row="1" FontSize="16"/>
        <TextBlock Name="CircularGrinding" HorizontalAlignment="Left" Margin="70,114,0,0" TextWrapping="Wrap" Text="Circular Grinding" VerticalAlignment="Top" Grid.Row="1" FontSize="16"/>
        <TextBlock Name="SurfaceGrinding" HorizontalAlignment="Left" Margin="247,88,0,0" TextWrapping="Wrap" Text="Surface Grinding" VerticalAlignment="Top" FontSize="16"/>
        <TextBlock Name="WireCutEDM" HorizontalAlignment="Left" Margin="247,114,0,0" TextWrapping="Wrap" Text="Wire Cut EDM" VerticalAlignment="Top" FontSize="16" Grid.RowSpan="2"/>
        <TextBlock Name="MillingFinish" HorizontalAlignment="Left" Margin="247,10,0,0" TextWrapping="Wrap" Text="Milling Finish" VerticalAlignment="Top" Grid.Row="1" FontSize="16"/>
        <TextBlock Name="TurningFinish" HorizontalAlignment="Left" Margin="247,36,0,0" TextWrapping="Wrap" Text="Turning Finish" VerticalAlignment="Top" Grid.Row="1" FontSize="16"/>
        <TextBlock Name="JigGrinding" HorizontalAlignment="Left" Margin="247,62,0,0" TextWrapping="Wrap" Text="Jig Grinding" VerticalAlignment="Top" Grid.Row="1" FontSize="16"/>
        <TextBlock Name="GrindingFinish" HorizontalAlignment="Left" Margin="247,88,0,0" TextWrapping="Wrap" Text="Grinding Finish" VerticalAlignment="Top" Grid.Row="1" FontSize="16"/>
        <TextBlock Name="ManualWorkstations" HorizontalAlignment="Left" Margin="247,114,0,0" TextWrapping="Wrap" Text="Manual Workstations" VerticalAlignment="Top" Grid.Row="1" FontSize="16"/>
        <TextBlock Name="QuantityText" HorizontalAlignment="Left" Margin="140,45,0,0" TextWrapping="Wrap" Text="Quantity" VerticalAlignment="Top" Height="22" Width="62" FontSize="16"/>
        
        <TextBox Name="XPPS" HorizontalAlignment="Left" Height="30" Margin="10,10,0,0" TextWrapping="NoWrap" Text="1450" VerticalAlignment="Top" Width="125" FontSize="18.667"/>
        <TextBox Name="PN" HorizontalAlignment="Left" Height="30" Margin="140,10,0,0" TextWrapping="NoWrap" Text="Part Name" VerticalAlignment="Top" Width="164" FontSize="18.667"/>
        <TextBox Name="DN" HorizontalAlignment="Left" Height="30" Margin="309,10,0,0" TextWrapping="NoWrap" Text="Drawing #" VerticalAlignment="Top" Width="129" FontSize="16"/>
        <TextBox Name="Quantity" HorizontalAlignment="Left" Height="22" Margin="78,45,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="57" RenderTransformOrigin="0.436,-0.358" FontSize="12" TextAlignment="Center"/>
        <TextBox Name="Material" HorizontalAlignment="Left" Height="22" Margin="207,45,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="38" RenderTransformOrigin="0.436,-0.358" FontSize="13.333" TextAlignment="Center"/>
        

        <TextBox Name="CAM" HorizontalAlignment="Left" Height="21" Margin="23,88,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="39" RenderTransformOrigin="0.436,-0.358" TextAlignment="Right"/>
        <TextBox Name="SAW" HorizontalAlignment="Left" Height="21" Margin="23,114,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="39" RenderTransformOrigin="0.436,-0.358" TextAlignment="Right" Grid.RowSpan="2"/>
        <TextBox Name="ST" HorizontalAlignment="Left" Height="21" Margin="23,10,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="39" RenderTransformOrigin="0.436,-0.358" TextAlignment="Right" Grid.Row="1"/>
        <TextBox Name="CN" HorizontalAlignment="Left" Height="21" Margin="23,36,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="39" RenderTransformOrigin="0.436,-0.358" TextAlignment="Right" Grid.Row="1"/>
        <TextBox Name="SM" HorizontalAlignment="Left" Height="21" Margin="23,62,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="39" RenderTransformOrigin="0.436,-0.358" TextAlignment="Right" Grid.Row="1"/>
        <TextBox Name="MAM" HorizontalAlignment="Left" Height="21" Margin="23,88,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="39" RenderTransformOrigin="0.436,-0.358" TextAlignment="Right" Grid.Row="1"/>
        <TextBox Name="CG" HorizontalAlignment="Left" Height="21" Margin="23,114,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="39" RenderTransformOrigin="0.436,-0.358" TextAlignment="Right" Grid.Row="1"/>
        <TextBox Name="SG" HorizontalAlignment="Left" Height="21" Margin="203,88,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="39" RenderTransformOrigin="0.436,-0.358" TextAlignment="Right"/>
        <TextBox Name="EDM" HorizontalAlignment="Left" Height="21" Margin="203,114,0,0" TextWrapping="Wrap" Text="0.0" VerticalAlignment="Top" Width="39" RenderTransformOrigin="0.436,-0.358" TextAlignment="Right" Grid.RowSpan="2"/>
        <TextBox Name="HM" HorizontalAlignment="Left" Height="21" Margin="203,10,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="39" RenderTransformOrigin="0.436,-0.358" TextAlignment="Right" Grid.Row="1"/>
        <TextBox Name="HT" HorizontalAlignment="Left" Height="21" Margin="203,36,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="39" RenderTransformOrigin="0.436,-0.358" TextAlignment="Right" Grid.Row="1"/>
        <TextBox Name="HMF" HorizontalAlignment="Left" Height="21" Margin="203,62,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="39" RenderTransformOrigin="0.436,-0.358" TextAlignment="Right" Grid.Row="1"/>
        <TextBox Name="HTF" HorizontalAlignment="Left" Height="21" Margin="203,88,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="39" RenderTransformOrigin="0.436,-0.358" TextAlignment="Right" Grid.Row="1"/>
        <TextBox Name="MW" HorizontalAlignment="Left" Height="21" Margin="203,114,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="39" RenderTransformOrigin="0.436,-0.358" TextAlignment="Right" Grid.Row="1"/>
        
        
        <TextBox Name="HRD" HorizontalAlignment="Left" Height="23" Margin="25,178,0,0" TextWrapping="Wrap" Text="Hardness" VerticalAlignment="Top" Width="106" FontSize="13.333" Grid.Row="1"/>
        <CheckBox Name="Heatreat" Content="Heat Treatment" HorizontalAlignment="Left" Margin="25,157,0,0" VerticalAlignment="Top" Width="103" Grid.Row="1"/>
        <CheckBox Name="COT" Content="Coating" HorizontalAlignment="Left" Margin="136,157,0,0" VerticalAlignment="Top" Grid.Row="1"/>
        <CheckBox Name="NIT" Content="Nitriding" HorizontalAlignment="Left" Margin="136,178,0,0" VerticalAlignment="Top" Grid.Row="1"/>
        <CheckBox Name="BOF" Content="Black Oxide Finish" HorizontalAlignment="Left" Margin="136,199,0,0" VerticalAlignment="Top" Grid.Row="1"/>
        <Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="72" Margin="16,150,0,0" VerticalAlignment="Top" Width="240" Grid.Row="1"/>

        <CheckBox Name="More" Content="Add more positions" HorizontalAlignment="Left" Margin="261,152,0,0" Grid.Row="1" VerticalAlignment="Top" Width="177" Height="23" FontSize="16"/>
        <Button Name="btnExit" Content="Continue" HorizontalAlignment="Left" Margin="261,180,0,0" VerticalAlignment="Top" Width="177" Grid.Row="1" Height="42" FontSize="24"/>
    </Grid>
</Window>
"@

#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; exit}

#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================

$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}


#===========================================================================
# Shows the form
#===========================================================================
$btnExit.Add_Click({
    if($XPPS.text.Length -ge 11 -and $PN.text -ne "Part Name" -and $Quantity.text -ne "Quantity" -and $DN.text -ne "Drawing #"){
        $check = $false
        $form.Close()
    }
})
$More.Add_Click({
    if($XPPS.text.Length -ge 11 -and $PN.text -ne "Part Name" -and $Quantity.text -ne "Quantity" -and $DN.text -ne "Drawing #"){
        $check = $true
        $form.Close()
    }
})

$Form.ShowDialog() | out-null
$checkbox = @($PN.text.trim(" "),$DN.text.trim(" "),$XPPS.text.trim(" "),$Material.text.trim(" "),$Quantity.text.trim(" "),$MAT.text.trim(" "),$CAM.text.trim(" "),$SAW.text.trim(" "),$ST.text.trim(" "),$CN.text.trim(" "),$SM.text.trim(" "),$MAM.text.trim(" "),$CG.text.trim(" "),$SG.text.trim(" "),$EDM.text.trim(" "),$HM.text.trim(" "),$HT.text.trim(" "),$HMF.text.trim(" "),$HTF.text.trim(" "),$MW.text.trim(" "),$Heatreat.IsChecked,$HRD.text.trim(" "),$NIT.IsChecked,$COT.IsChecked,$BOF.IsChecked,$More.IsChecked)

return $checkbox
}

$XPPS1 = 1450
for($i=0;$i -lt 3000;$i++){
$checkbox += ,(CheckBox $xpps1)
$XPPS1 = $checkbox[$i][2]
if(!$checkbox[$i][26]){break}
}



<# Removed crap



        <TextBox Name="textBox3_Copy16" HorizontalAlignment="Left" Height="22" Margin="384,45,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="54" RenderTransformOrigin="0.436,-0.358" TextAlignment="Center" Text="XXX" FontSize="13.333"/>



        <TextBox Name="textBox3_Copy15" HorizontalAlignment="Left" Height="22" Margin="294,45,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="54" RenderTransformOrigin="0.436,-0.358" FontSize="13.333" TextAlignment="Center" Text="290"/>

        <TextBlock Name="XVAS" HorizontalAlignment="Left" Margin="250,45,0,0" TextWrapping="Wrap" Text="XVAS" VerticalAlignment="Top" Height="22" Width="39" FontSize="16"/>
        <TextBlock Name="PO" HorizontalAlignment="Left" Margin="353,45,0,0" TextWrapping="Wrap" Text="PO" VerticalAlignment="Top" Height="22" Width="26" FontSize="16"/>



#>