
[boolean]$script:initialized = $false
[string]$script:outPath = Join-Path $env:TEMP 'SampleData'

function Start-SampleData
{
    [CmdletBinding()]param()

    if(! $script:initialized)
    {
        Write-Verbose "Start-SampleData"

        if( $null -eq $script:word)
        {
            Write-Verbose "Create Word Object"
            Add-Type -AssemblyName "Microsoft.Office.Interop.Word" | Out-Null
            $script:word = New-Object -ComObject word.application
        }

        if( $null -eq $script:companies)
        {
            $companiesFilePath = Join-Path $PSScriptRoot 'companies.csv'
            Write-Verbose "Get Companies Data ($companiesFilePath.csv)"
            $script:companies = Import-Csv $companiesFilePath
        }

        if( $null -eq $script:rand)
        {
            Write-Verbose "Create Rand Object"
            $script:rand = New-Object -Type System.Random
        }

        Write-Verbose "Output path: $script:outPath"
        if( ! (Test-Path $script:outPath) )
        {
            Write-Verbose "Create Output Dir: $script:outPath"
            New-Item -Path $script:outPath -ItemType Directory | Out-Null
        }

        $script:initialized = $true
    }
}


function Stop-SampleData
{
    [CmdletBinding()]param()

    try
    {
        Write-Verbose "Stop-SampleData"
        $word.Quit()
    }
    catch 
    {
        Write-Warning $_
    }

    $script:initialized = $false
}


function Get-InvoiceData
{
    [CmdletBinding()]param()

    Write-Verbose "Get-InvoiceData"

    if(!$script:initialized)
    {
        Start-SampleData
       #throw 'Not Initialized. Call Start-SampleData'
    }

    $index = $script:rand.Next(0,($companies.Count))
    $invoiceData = @{}
    $invoiceData['Company'] = "{0} ({1})" -f $companies[$index].COMPANY.Trim('.'), $companies[$index].COUNTRY
    $invoiceData['InvoiceDate'] = Get-Date -Year (Get-Random -Minimum 2001 -Maximum 2018) -Month (Get-Random -Minimum 1 -Maximum 12) -Day (Get-Random -Minimum 1 -Maximum 28)
    $invoiceData['Filename'] = "Invoice - {0} - {1}.pdf" -f $invoiceData['Company'], ($invoiceData['InvoiceDate'].ToString('yyyy-MM-dd'))
    $invoiceData['Value'] = Get-Random -min 1000 -max 100000

    $invoicedata = New-Object -TypeName PSObject -Property $InvoiceData 
    Write-Output $invoiceData
}

function New-InvoiceFile
{
    [CmdletBinding()]param(
        [Parameter(Mandatory=$false)]
        [string]$OutputFolder
    )

    if([string]::IsNullOrEmpty($OutputFolder))
    {
        $OutputFolder = $script:outPath
    }

    $InvoiceData = Get-InvoiceData
    $invoiceData['FilePath'] = Join-Path $OutputFolder $invoiceData['Filename'] 

    Write-Verbose "Create Document in Word"
    $doc = $script:word.Documents.add()
    $selection = $Word.Selection
    $selection.Text = $invoiceData.Filename

    [string]$filePath = $InvoiceData.FilePath
    Write-Verbose "Save: $filePath"
    $fmt = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatPDF
    $doc.SaveAs( [ref]$filePath, [ref]$fmt )
    $doc.Close( [ref]([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges) )

    Write-Verbose 'Update File Date'
    Set-ItemProperty -Path $filePath -Name CreationTime  -Value $invoiceData.InvoiceDate
    Set-ItemProperty -Path $filePath -Name LastWriteTime -Value $invoiceData.InvoiceDate

    Write-Output $InvoiceData
}


function New-InvoiceFiles
{
    [CmdletBinding()]param(
        [int]$Count
    )

    1..$Count | %{ New-InvoiceFile }

}


function Clear-SampleDataFiles
{
    [CmdletBinding()]param()

    Write-Verbose "Clear-SampleDataFiles"

    try
    {
        Remove-Item -Path $script:outPath -Recurse -Force
    }
    catch
    {
        Write-Verbose $_
    }
}



function Get-HireDate
{
    $dt = [DateTime](Get-Random -Minimum (Get-Date '1/1/1990').Ticks  -Maximum (Get-Date '1/1/2010').Ticks)
#   $dt = [DateTime](Get-Random -Minimum (Get-Date '1/1/1990').Ticks  -Maximum (Get-Date).Ticks)
#   $dt.ToShortDateString()
    $dt
}


function Get-TermDate([DateTime]$HireDate)
{
    if( [Boolean](Get-Random -Minimum 0 -Maximum 2) ) { 
        $dt = [DateTime](Get-Random -Minimum $HireDate.Ticks -Maximum (Get-Date).Ticks)
#       $dt.ToShortDateString()
        $dt
    } 
}



function Get-RandomFirstName
{
    [CmdletBinding()]
    Param(
        $basepath = $PSScriptRoot # (Get-Module EmployeeFiles).ModuleBase
    )

    if($null -eq $script:rand)
    {
        Write-Verbose 'New Random'
        $script:rand = New-Object -Type System.Random
    }
    
    if($null -eq $script:firstnames) 
    {
        Write-Verbose (Join-Path $basepath 'firstnames.csv')
        $script:firstnames = Import-Csv ( Join-Path $basepath 'firstnames.csv' )
    }

    $script:firstnames[$rand.Next(0,($firstnames.Count))].Name
}
#Export-ModuleMember Get-RandomFirstName


function Get-RandomLastName
{
    [CmdletBinding()]
    Param(
        $basepath = $PSScriptRoot #(Get-Module EmployeeFiles).ModuleBase
    )

    if($null -eq $script:rand)
    {
        Write-Verbose 'New Random'
        $script:rand = New-Object -Type System.Random
    }
    
    if($null -eq $script:lastnames) 
    {
        Write-Verbose (Join-Path $basepath 'lastnames.csv')
        $script:lastnames = Import-Csv ( Join-Path $basepath 'lastnames.csv' )
    }

    $script:lastnames[$rand.Next(0,($lastnames.Count))].Name
}
#Export-ModuleMember Get-RandomLastName


function Get-RandomEmployee
{
    [CmdletBinding()]
    Param(
    )

    $first = Get-RandomFirstName
    $last  = Get-RandomLastName
    $empId = Get-Random -min 100000 -max 999999
    $hireDate = Get-HireDate
    $termDate = Get-TermDate -HireDate $hireDate

    $file  = [PSCustomObject][ordered]@{
        'Employee Level'   = if ( (Get-Random)%3 -eq 0) {'Executive'} else { "General" } 
        'First Name'       = $first
        'Last Name'        = $last
        'Employee Name'    = "$first $last"
        'Employee Id'      = $empId
        'Hire Date'        = $hireDate
        'Termination Date' = $termDate
    }

    $file
}
#Export-ModuleMember -Function Get-RandomEmployee



function Get-RandomEmployeeFile
{
    [CmdletBinding()]
    param(
        [string]$Container,
        [string]$Category
    )

    $file = Get-RandomEmployee

    $profile = 'Employee File'
    $title = "$($profile): $($file.'Employee Name') ($($file.'Employee Id'))"

    Add-Member -InputObject $file -Name 'Title' -Value $title -MemberType NoteProperty
    Add-Member -InputObject $file -Name 'Profile' -Value $profile -MemberType NoteProperty

    if($null -ne $Container)
    {
        Add-Member -InputObject $file -Name 'Container' -Value $Container -MemberType NoteProperty
    }

    if($null -ne $Category)
    {
        Add-Member -InputObject $file -Name 'Category' -Value $Category -MemberType NoteProperty
    }

    $file
}
#Export-ModuleMember -Function Get-RandomEmployeeFile


function Get-RandomEmployeeFilesImportFile
{
    [CmdletBinding()]
    param(
        [string]$Path,

        [string]$Container,

        [string]$Category,

        [int]$Count=1,

        [switch]$Overwrite
    )

    if( ($null -ne $Path) -and ($Path -ne [string]::Empty) )
    {
        if( (Test-Path $Path) -and !$Overwrite )
        {
            Write-Error "Output File Already Exists, use Overwrite switch to overwrite the file."
            return
        }
    }

    $filesData = @()
    for($i=0; $i -lt $Count; $i++)
    {
        $file = Get-RandomEmployeeFile -Container $Container -Category $Category
        $file #send $file to the pipeline        
        Write-Verbose $file

        $filesData += $file

    }

    if( ($null -ne $Path) -and ($Path -ne [string]::Empty) )
    {
        $filesData | Export-Csv $Path -NoTypeInformation -Force
        Write-Verbose "Exported Employee Files to '$Path'"
    }

}
#Export-ModuleMember -Function Get-RandomEmployeeFilesImportFile

