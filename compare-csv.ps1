# The comments are just for me to change one or more things if I will need them
# Of course you can use the cmd way and and delete the dialog boxes

Add-Type -AssemblyName System.Windows.Forms
function Open-DialogBox {

    Param (
    [Parameter(Mandatory=$False)]
    [string] $InitialDirectory #= [Environment]::GetFolderPath('Desktop')
    )
    
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Select CSV file"
    $OpenFileDialog.InitialDirectory = $InitialDirectory
    $OpenFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName

}

function Save-DialogBox {

    Param (
    [Parameter(Mandatory=$False)]
    [string] $InitialDirectory
    )
    
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Title = "Select csv file to save or create a new one"
    $SaveFileDialog.InitialDirectory = $InitialDirectory
    $SaveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $SaveFileDialog.OverwritePrompt = $True;
    $SaveFileDialog.ShowDialog() | Out-Null
    return $SaveFileDialog.FileName
}

function Show-Toast {

    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning 
    $balloon.BalloonTipText = "No file will be saved!"
    $balloon.BalloonTipTitle = "Hallo $Env:USERNAME! File Error!" 
    $balloon.Visible = $true 
    $balloon.ShowBalloonTip(5000)

}


#$file1 = Read-Host "full path to your first csv file (like C:\teams\team.csv)"
#Write-Host -ForegroundColor Yellow $file1
$file1 = Open-DialogBox
$data1 = Import-Csv -Path $file1 -Delimiter ";" #| Select-Object -First 1
#$data1 | Format-Table

#$file2 = Read-Host "full path to your second csv file (like C:\teams\team.csv)"
#Write-Host -ForegroundColor Yellow $file2
$file2 = Open-DialogBox
$data2 = Import-Csv -Path $file2 -Delimiter ";" #| Select-Object -First 1
#$data2 | Format-Table

#$export = Read-Host "where to export the new file (like C:\teams\team.csv)"
#$export = Save-DialogBox



# Properties to compare
$commonProps1 = [Linq.Enumerable]::Intersect([string[]] $data1[0].PSObject.Properties.Name, [string[]] $data2[0].PSObject.Properties.Name)
$compareproperty = $commonProps1 | Out-GridView -OutputMode 'Multiple' -Title 'Select Property Names to Compare'


# Properties to write in new file
$commonProps2 = [Linq.Enumerable]::Intersect([string[]] $data1[0].PSObject.Properties.Name, [string[]] $data2[0].PSObject.Properties.Name)
$outputobject = $commonProps2 | Out-GridView -OutputMode 'Multiple' -Title 'Select Property Names for output in new csv file'




<####   This is the part if you want to read the properties from command line

#$compareproperty = Read-Host "which column do you want to compare?"

######################################################################################################################
[System.Collections.ArrayList]$comparecolumns = @()
do {
 $Columns = (Read-Host "Please enter the column names to compare from csv, one after one with return. To abort leave empty and press return")
 if ($Columns -ne '') {$comparecolumns += $Columns}
}
until ($Columns -eq '')

if($comparecolumns.count -eq 1){
    $compareproperty = $outputcolumns
    } else {
    $compareproperty = $outputcolumns -join ", "
}
######################################################################################################################

######################################################################################################################
[System.Collections.ArrayList]$outputcolumns = @()
do {
 $Columns = (Read-Host "Please enter the column names to output in csv, one after one with return. To abort leave empty and press return")
 if ($Columns -ne '') {$outputcolumns += $Columns}
}
until ($Columns -eq '')

if($outputcolumns.count -eq 1){
    $outputobject = $outputcolumns
    } else {
    $outputobject = $outputcolumns -join ", "
}
######################################################################################################################

####>



<### This is to show which entries has the same values but it only works with one property 

#Compare both CSV files
$Results = Compare-Object  $data1 $data2 -Property $compareproperty -IncludeEqual
 
$Array = @()       
Foreach($R in $Results)
{
    If( $R.sideindicator -eq "==" )
    {
        $Object = [pscustomobject][ordered] @{
 
            Columns = $R.$compareproperty
            "Compare indicator" = $R.sideindicator
 
        }
        $Array += $Object
    }
}
 
#Count users in both files
($Array | sort-object $compareproperty | Select-Object * -Unique).count
 
#Display results in console
$Array | Out-GridView
###>


$SaveFile = Save-DialogBox

if($SaveFile -ne ""){
    Compare-Object $data1 $data2 -Property $compareproperty | Select-Object $outputobject | export-csv -Path $SaveFile -Delimiter ";" -NoTypeInformation
} else {
     #[System.Windows.Forms.MessageBox]::Show("ATTENTION! $([System.Environment]::NewLine)No file will be saved!","File Error!",3,[System.Windows.Forms.MessageBoxIcon]::Exclamation)
     Show-Toast
}

#Compare-Object $data1 $data2 -Property $compareproperty | Select-Object $outputobject | export-csv $export -Delimiter ";" -NoTypeInformation
#$export | Out-GridView
