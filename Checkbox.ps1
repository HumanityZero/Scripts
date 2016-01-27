Function CheckBox{
#==============================================================================================
# XAML Code - Imported from Visual Studio Express WPF Application
#==============================================================================================
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="New Part" Height="350" Width="447.911">
    <Grid Margin="0,0,-0.2,0.2">
        <Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="72" Margin="48,231,0,0" VerticalAlignment="Top" Width="240"/>
        
        <CheckBox Name="MAT" Content="Material" HorizontalAlignment="Left" Margin="10,98,0,0" VerticalAlignment="Top" IsChecked="True"/>
        <CheckBox Name="SAW" Content="Sawing" HorizontalAlignment="Left" Margin="75,98,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="CAM" Content="CAM" HorizontalAlignment="Left" Margin="10,147,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="ST" Content="Standard Turning" HorizontalAlignment="Left" Margin="75,119,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="CN" Content="CNC Turning" HorizontalAlignment="Left" Margin="75,140,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="SM" Content="Standard Milling" HorizontalAlignment="Left" Margin="75,161,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="MAM" Content="Multiaxial Milling" HorizontalAlignment="Left" Margin="75,182,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="CG" Content="Circular grinding" HorizontalAlignment="Left" Margin="75,203,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="SG" Content="Surface Grinding" HorizontalAlignment="Left" Margin="189,98,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="EDM" Content="Wire Cut EDM" HorizontalAlignment="Left" Margin="189,119,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="HM" Content="Hard Milling" HorizontalAlignment="Left" Margin="189,140,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="HT" Content="Hard Turning" HorizontalAlignment="Left" Margin="189,161,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="HMF" Content="Hard Milling Finish" HorizontalAlignment="Left" Margin="189,182,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="HTF" Content="Hard Turning Finsh" HorizontalAlignment="Left" Margin="189,203,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="MW" Content="Manual Workstations" HorizontalAlignment="Left" Margin="299,98,0,0" VerticalAlignment="Top" IsChecked="True"/>
        
        <Button Name="btnExit" Content="Ok" HorizontalAlignment="Left" Margin="353,267,0,0" VerticalAlignment="Top" Width="75"/>
        
        <CheckBox Name="NIT" Content="Nitriding" HorizontalAlignment="Left" Margin="168,259,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="COT" Content="Coating" HorizontalAlignment="Left" Margin="168,238,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="Heatreat" Content="Heat Treatment" HorizontalAlignment="Left" Margin="57,238,0,0" VerticalAlignment="Top" Width="103"/>
        <CheckBox Name="BOF" Content="Black Oxid Finish" HorizontalAlignment="Left" Margin="168,280,0,0" VerticalAlignment="Top"/>
        
        <TextBox Name="PN" HorizontalAlignment="Left" Height="38" Margin="10,10,0,0" TextWrapping="Wrap" Text="Part Name" VerticalAlignment="Top" Width="341" FontSize="18.667"/>
        <TextBox Name="Quantity" HorizontalAlignment="Left" Height="38" Margin="356,10,0,0" TextWrapping="Wrap" Text="Quantity" VerticalAlignment="Top" Width="74" FontSize="16"/>
        <TextBox Name="DN" HorizontalAlignment="Left" Height="30" Margin="10,53,0,0" TextWrapping="Wrap" Text="Drawing Number" VerticalAlignment="Top" Width="420" AutoWordSelection="True" FontSize="16"/>
        <TextBox Name="Material" HorizontalAlignment="Left" Height="23" Margin="10,119,0,0" TextWrapping="Wrap" Text="Material" VerticalAlignment="Top" Width="60"/>
        <TextBox Name="HRD" HorizontalAlignment="Left" Height="23" Margin="57,259,0,0" TextWrapping="Wrap" Text="Hardness" VerticalAlignment="Top" Width="106" FontSize="13.333"/>
        <TextBox Name="XPPS" HorizontalAlignment="Left" Height="28" Margin="338,234,0,0" TextWrapping="Wrap" Text="$XPPS1" VerticalAlignment="Top" Width="90" FontSize="13.333"/>
        <CheckBox Name="Print" Content="Print?" HorizontalAlignment="Left" Margin="376,294,0,0" VerticalAlignment="Top" Width="52" IsChecked="True"/>
        <CheckBox Name="More" Content="More" HorizontalAlignment="Left" Margin="322,294,0,0" VerticalAlignment="Top" Width="49" IsChecked="True"/>

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
    if($XPPS.text.Length -ge 11 -and $PN.text -ne "Part Name" -and $Quantity.text -ne "Quantity" -and $DN.text -ne "Drawing Number"){
        $form.Close()
    }
})
$Form.ShowDialog() | out-null
$checkbox = @($PN.text,$DN.text,$XPPS.text,$Material.text,$Quantity.text,$Print.IsChecked,$Mat.IsChecked,$CAM.IsChecked,$SAW.IsChecked,$ST.IsChecked,$CN.IsChecked,$SM.IsChecked,$MAM.IsChecked,$CG.IsChecked,$SG.IsChecked,$EDM.IsChecked,$HM.IsChecked,$HT.IsChecked,$HMF.IsChecked,$HTF.IsChecked,$MW.IsChecked,$Heatreat.IsChecked,$HRD.text,$NIT.IsChecked,$COT.IsChecked,$BOF.IsChecked,$More.IsChecked)

return $checkbox
}

$XPPS1 = 1450
for($i=0;$i -lt 3000;$i++){
$checkbox += ,(CheckBox $xpps1)
$XPPS1 = $checkbox[$i][2]
if(!$checkbox[$i][26]){break}
}