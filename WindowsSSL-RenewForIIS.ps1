# Assisted method of renewing SSL certificate on an IIS server via a simple GUI
# Must be run on a web server

#========#
# ^^^^^^ #
# README #
#========#

########################################################################################################################################################################################################################

Add-Type -AssemblyName PresentationFramework

$CerFile = Test-Path -Path C:\Parceu\IISYSSL.cer
if ($CerFile -eq $true)
{
    Remove-Item C:\Parceu\IISYSSL.cer -Force | Out-Null
}
$InfFile = Test-Path -Path C:\Parceu\IISYSSL.inf
if ($InfFile -eq $true)
{
    Remove-Item C:\Parceu\IISYSSL.inf -Force | Out-Null
}
$csrFile = Test-Path -Path C:\Parceu\IISYSSL.csr
if ($csrFile -eq $true)
{
    Remove-Item C:\Parceu\IISYSSL.csr -Force | Out-Null
}

Function New-SSLCert
{
    [CmdletBinding()]
    param ([Parameter(Mandatory)][string]$CertDomain)

    $ParceuCFolder = Test-Path -Path C:\Parceu\
    if ($ParceuCFolder -ne $true)
    {
        New-Item -ItemType "directory" -Path "C:\Parceu" | Out-Null
    }
    $CerFile = Test-Path -Path C:\Parceu\IISYSSL.cer
    if ($CerFile -eq $true)
    {
        Remove-Item C:\Parceu\IISYSSL.cer -Force | Out-Null
    }
    $InfFile = Test-Path -Path C:\Parceu\IISYSSL.inf
    if ($InfFile -eq $true)
    {
        Remove-Item C:\Parceu\IISYSSL.inf -Force | Out-Null
    }
    $csrFile = Test-Path -Path C:\Parceu\IISYSSL.csr
    if ($csrFile -eq $true)
    {
        Remove-Item C:\Parceu\IISYSSL.csr -Force | Out-Null
    }
$INFVerSignature = '$Windows NT$'

$INFFileContent =
@"
[Version]
Signature = "$INFVerSignature"
[NewRequest]
Subject = "CN=$CertDomain"
Exportable = TRUE
KeyLength = 2048
KeySpec = 1
KeyUsage = 0xa0
MachineKeySet = True
ProviderName = "Microsoft RSA SChannel Cryptographic Provider"
ProviderType = 12
Silent = True
SMIME = False
RequestType = PKCS10
"@

$INFFileContent | out-file -filepath C:\Parceu\IISYSSL.inf -force
certreq.exe -new C:\Parceu\IISYSSL.inf C:\Parceu\IISYSSL.csr
$var_txtCSR.Text = get-content C:\Parceu\IISYSSL.csr

}

Function Add-SSLCert
{
    $var_inputCER.Text | Out-File -FilePath C:\Parceu\IISYSSL.cer -force
    certreq.exe -accept -machine C:\Parceu\IISYSSL.cer
    $CerFile = Test-Path -Path C:\Parceu\IISYSSL.cer
    if ($CerFile -eq $true)
    {
        Remove-Item C:\Parceu\IISYSSL.cer -Force | Out-Null
    }
    $InfFile = Test-Path -Path C:\Parceu\IISYSSL.inf
    if ($InfFile -eq $true)
    {
        Remove-Item C:\Parceu\IISYSSL.inf -Force | Out-Null
    }
    $csrFile = Test-Path -Path C:\Parceu\IISYSSL.csr
    if ($csrFile -eq $true)
    {
        Remove-Item C:\Parceu\IISYSSL.csr -Force | Out-Null
    }

}

$inputXML = @"
<Window x:Name="IISYSSL" x:Class="IISYSSL.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:IISYSSL"
        mc:Ignorable="d"
        Title="IISYSSL" Height="418.499" Width="490.894" ResizeMode="CanMinimize">
    <Grid Margin="0,0,4,0">
        <TextBox x:Name="txtCSR" Margin="10,0,0,60" TextWrapping="Wrap" Text="" IsReadOnly="True" HorizontalAlignment="Left" Width="223" Height="231" VerticalAlignment="Bottom"/>
        <TextBox x:Name="InputCER" Margin="243,0,0,60" TextWrapping="Wrap" Text="" HorizontalAlignment="Left" Width="223" Height="231" VerticalAlignment="Bottom" AcceptsReturn="True"/>
        <Button x:Name="btnCOPY" Content="Copy CSR" Margin="88,332,0,10" HorizontalAlignment="Left" Width="71"/>
        <Button x:Name="btnIMPORT" Content="Import CER" Margin="313,332,0,10" HorizontalAlignment="Left" Width="71"/>
        <TextBox x:Name="inputDOMAIN" HorizontalAlignment="Left" Margin="13,24,0,341" TextWrapping="Wrap" Text="" Width="456"/>
        <Button x:Name="btnREQUEST" Content="Request CSR" Margin="175,52,0,296" HorizontalAlignment="Left" Width="122"/>
        <Label x:Name="DomainLabel" Content="Domain" HorizontalAlignment="Left" Margin="7,2,0,0" VerticalAlignment="Top" Height="27" Width="150"/>
        <Label x:Name="CSRLabel" Content="CSR" HorizontalAlignment="Center" Margin="13,67,432,0" VerticalAlignment="Top" Height="27" Width="36"/>
        <Label x:Name="CERLabel" Content="CER" HorizontalAlignment="Center" Margin="440,69,6,0" VerticalAlignment="Top" Height="27" Width="35"/>
        <Label Content="By packet" HorizontalAlignment="Left" Margin="201,353,0,0" VerticalAlignment="Top" Width="77"/>
        <Label Content="v2.0" HorizontalAlignment="Left" Height="26" Margin="433,353,0,0" VerticalAlignment="Top" Width="38"/>
    </Grid>
</Window>
"@

$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
[XML]$XAML = $inputXML

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load( $reader )
} catch {
    Write-Warning $_.Exception
    throw
}
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
    } catch {
        throw
    }
}

$var_btnREQUEST.Add_Click( 
{
    $var_txtCSR.Text = ""
    New-SSLCert -CertDomain $var_inputDOMAIN.Text
})

$var_btnCOPY.Add_Click( 
{
    get-content C:\Parceu\IISYSSL.csr | clip
    [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [System.Windows.Forms.MessageBox]::Show('CSR copied to clipboard','Notice')
})

$var_btnIMPORT.Add_Click( 
{
    Add-SSLCert
    $var_txtCSR.Text = ""
    $var_inputDOMAIN.Text = ""
    $var_inputCER.Text = ""
})

$Null = $window.ShowDialog()
