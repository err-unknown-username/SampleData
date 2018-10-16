
function Initialize-Data
{
    $companiesFilePath = Join-Path $PSScriptRoot 'companies.csv'
    Write-Verbose "Get Companies Data: $companiesFilePath" -Verbose
    $script:companies = Import-Csv $companiesFilePath

    $firstnamesPath = Join-Path $PSScriptRoot 'firstnames.csv'
    Write-Verbose "Get Firstname Data: $firstnamesPath" -Verbose
    $script:firstnames = Import-Csv $firstnamesPath

    $lastnamesPath = Join-Path $PSScriptRoot 'lastnames.csv'
    Write-Verbose "Get Lastname Data:  $lastnamesPath" -Verbose
    $script:lastnames = Import-Csv $lastnamesPath

    Write-Verbose "Create Rand Object"
    $script:rand = New-Object -Type System.Random
}
Initialize-Data

function New-Document
{
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Filename,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('pdf','docx','doc','txt','rtf')]
        [string]$FileType = 'pdf',

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$FileContent,

        [Parameter(Mandatory=$true,ParameterSetName='LiteralPath')]
        [ValidateNotNullOrEmpty()]
        [string]$OutputFolder = ( Join-Path ([Environment]::GetFolderPath('Desktop')) 'Samples' ),

        [Parameter(Mandatory=$true,ParameterSetName='RelativePath')]
        [ValidateNotNullOrEmpty()]
        [string]$OutputFolderRel,

        [Parameter(Mandatory=$false)]
        [datetime]$CreationTime = (Get-Date),

        [Parameter(Mandatory=$false)]
        [datetime]$LastWriteTime = (Get-Date),

        [Parameter(Mandatory=$false)]
        [switch]$Force
    )

    BEGIN
    {
        $wordObject = New-Object -ComObject word.application
    }

    PROCESS
    {
        if( $PSCmdlet.ParameterSetName -eq 'RelativePath' )
        {
            $OutputFolder = Join-Path $OutputFolder $OutputFolderRel
        }

        if( !(Test-Path $OutputFolder) )
        {
            if( $Force.IsPresent )
            {
                New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
            }
            else
            {
                throw "OutputFolder Path does not exist. Use '-Force' to create. '$OutputFolder'"
            }
        }

        [string]$filePath = Join-Path $OutputFolder $Filename
        Write-Verbose "Save: $filePath"

        Write-Verbose "Create Document in Word"
        $doc = $wordObject.Documents.add()
        $selection = $wordObject.Selection
        if( [string]::IsNullOrEmpty($FileContent) )
        {
            $selection.Text = $Filename
        }
        else
        {
            $selection.Text = $FileContent
        }

        switch($FileType) {
            'pdf'  { $fmt = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatPDF; break }
            'docx' { $fmt = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault; break }
            'doc'  { $fmt = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocument; break }
            'txt'  { $fmt = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatText; break }
            'rtf'  { $fmt = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatRTF; break }
        }
        $doc.SaveAs( [ref]$filePath, [ref]$fmt )
        $doc.Close( [ref]([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges) )

        Write-Verbose 'Update File Date'
        Set-ItemProperty -Path $filePath -Name CreationTime  -Value $CreationTime
        Set-ItemProperty -Path $filePath -Name LastWriteTime -Value $LastWriteTime

        Write-Output $filePath
    }

    END
    {
        $wordObject.Quit()
    }

}

function Get-InvoiceData
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [ValidateSet('pdf','docx','doc','txt','rtf')]
        [string]$FileType = 'pdf'
    )

    Write-Verbose "Get-InvoiceData"

    $index = $script:rand.Next(0,($companies.Count))
    $invoiceData = @{}
    $invoiceData['Company'] = "{0} ({1})" -f $companies[$index].COMPANY.Trim('.'), $companies[$index].COUNTRY
    $invoiceData['InvoiceDate'] = Get-Date -Year (Get-Random -Minimum 2001 -Maximum 2018) -Month (Get-Random -Minimum 1 -Maximum 12) -Day (Get-Random -Minimum 1 -Maximum 28)
    $invoiceData['Filename'] = "Invoice - {0} - {1}.{2}" -f $invoiceData['Company'], ($invoiceData['InvoiceDate'].ToString('yyyy-MM-dd')), $FileType
    $invoiceData['Value'] = Get-Random -min 1000 -max 100000

    $invoicedata = New-Object -TypeName PSObject -Property $InvoiceData 
    Write-Output $invoiceData
}

function Get-CompanyData
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [ValidateSet('pdf','docx','doc','txt','rtf')]
        [string]$FileType = 'pdf'
    )

    Write-Verbose "Get-CompanyData"

    $index = $script:rand.Next(0,($companies.Count))
    $companyData = @{}
    $companyData['Company']      = $companies[$index].COMPANY.Trim('.')
    $companyData['Country']      = $companies[$index].COUNTRY
    $companyData['FileType']     = $FileType
    $companyData['CompanyDate']  = Get-Date -Year (Get-Random -Minimum 2001 -Maximum 2018) -Month (Get-Random -Minimum 1 -Maximum 12) -Day (Get-Random -Minimum 1 -Maximum 28)
    $companyData['CompanyValue'] = Get-Random -min 1000 -max 100000
    $companyData['Filename']     = "{0} - {1}.{2}" -f $companyData['Company'], ($companyData['CompanyDate'].ToString('yyyy-MM-dd')), $FileType

    $companyData = New-Object -TypeName PSObject -Property $companyData 
    Write-Output $companyData
}

function New-InvoiceFile
{
    [CmdletBinding(DefaultParameterSetName='LiteralPath')]
    param(
        [Parameter(Mandatory=$false,ParameterSetName='LiteralPath')]
        [ValidateNotNullOrEmpty()]
        [string]$OutputFolder = ( Join-Path ([Environment]::GetFolderPath('Desktop')) 'Samples' ),

        [Parameter(Mandatory=$false,ParameterSetName='RelativePath')]
        [ValidateNotNullOrEmpty()]
        [string]$OutputFolderRel = 'Invoices',

        [Parameter(Mandatory=$false)]
        [ValidateSet('pdf','docx','doc','txt','rtf')]
        [string]$FileType = 'pdf',

        [Parameter(Mandatory=$false)]
        [switch]$Force
    )

    $InvoiceData = Get-InvoiceData -FileType $FileType

    $params = @{
        Filename = $invoiceData.Filename
        FileContent = ($InvoiceData | ConvertTo-Json)
        CreationTime = $InvoiceData.InvoiceDate
        LastWriteTime = $InvoiceData.InvoiceDate
        FileType = $FileType
        Force = $Force.IsPresent
    }

    if($PSCmdlet.ParameterSetName -eq 'LiteralPath') {
        $params['OutputFolder'] = $OutputFolder
    } 
    else {
        $params['OutputFolderRel'] = $OutputFolderRel
    }

    New-Document @params
}


function New-InvoiceFiles
{
    [CmdletBinding(DefaultParameterSetName='LiteralPath')]
    param(
        [Parameter(Mandatory=$false)]
        [int]$Count = 10,

        [Parameter(Mandatory=$false,ParameterSetName='LiteralPath')]
        [ValidateNotNullOrEmpty()]
        [string]$OutputFolder = ( Join-Path ([Environment]::GetFolderPath('Desktop')) 'Samples' ),

        [Parameter(Mandatory=$false,ParameterSetName='RelativePath')]
        [ValidateNotNullOrEmpty()]
        [string]$OutputFolderRel = 'Invoices',

        [Parameter(Mandatory=$false)]
        [ValidateSet('pdf','docx','doc','txt','rtf')]
        [string]$FileType = 'pdf'

    )
    

    1..$Count | %{ 
        if($PSCmdlet.ParameterSetName -eq 'LiteralPath') {
            New-InvoiceFile -OutputFolder $OutputFolder -FileType $FileType 
        } 
        else {
            New-InvoiceFile -OutputFolderRel $OutputFolderRel -FileType $FileType 
        }
    }

}


function New-ContractFile
{
    [CmdletBinding(DefaultParameterSetName='LiteralPath')]
    param(
        [Parameter(Mandatory=$false,ParameterSetName='LiteralPath')]
        [ValidateNotNullOrEmpty()]
        [string]$OutputFolder = ( Join-Path ([Environment]::GetFolderPath('Desktop')) 'Samples' ),

        [Parameter(Mandatory=$false,ParameterSetName='RelativePath')]
        [ValidateNotNullOrEmpty()]
        [string]$OutputFolderRel = 'Invoices',

        [Parameter(Mandatory=$false)]
        [ValidateSet('pdf','docx','doc','txt','rtf')]
        [string]$FileType = 'pdf'
    )

    $InvoiceData = Get-InvoiceData -FileType $FileType

    $contractName = "{0} - {1}" -f $InvoiceData.InvoiceDate.Year, $InvoiceData.Company
    @('SoW','Contract','Appendix1','Appendix2','Appendix2') | %{
        $params = @{
            Filename = "{0} - {1}.{2}" -f $contractName, $_, $FileType
            FileContent = ($InvoiceData | ConvertTo-Json)
            CreationTime = $InvoiceData.InvoiceDate
            LastWriteTime = $InvoiceData.InvoiceDate
            FileType = $FileType
        }

        if($PSCmdlet.ParameterSetName -eq 'LiteralPath') {
            $params['OutputFolder'] = Join-Path $OutputFolder $contractName
        } 
        else {
            $params['OutputFolderRel'] = "{0}\{1}" -f $OutputFolderRel, $contractName
        }

        New-Document @params -Force
    }
}


function New-ContractFiles
{
    [CmdletBinding(DefaultParameterSetName='LiteralPath')]
    param(
        [Parameter(Mandatory=$false)]
        [int]$Count = 10,

        [Parameter(Mandatory=$false,ParameterSetName='LiteralPath')]
        [ValidateNotNullOrEmpty()]
        [string]$OutputFolder = ( Join-Path ([Environment]::GetFolderPath('Desktop')) 'Samples' ),

        [Parameter(Mandatory=$false,ParameterSetName='RelativePath')]
        [ValidateNotNullOrEmpty()]
        [string]$OutputFolderRel = 'Invoices',

        [Parameter(Mandatory=$false)]
        [ValidateSet('pdf','docx','doc','txt','rtf')]
        [string]$FileType = 'pdf'

    )
    

    1..$Count | %{ 
        if($PSCmdlet.ParameterSetName -eq 'LiteralPath') {
            New-ContractFile -OutputFolder $OutputFolder -FileType $FileType 
        } 
        else {
            New-ContractFile -OutputFolderRel $OutputFolderRel -FileType $FileType 
        }
    }

}



function Get-HireDate
{
    param(
        [Parameter(Mandatory=$false)]
        [datetime]$MinHireDate = (Get-Date '1/1/1990'),

        [Parameter(Mandatory=$false)]
        [datetime]$MaxHireDate = (Get-Date '1/1/2010')
    )

    $dt = [DateTime](Get-Random -Minimum ($MinHireDate).Ticks  -Maximum ($MaxHireDate).Ticks)
    Write-Output $dt
}


function Get-TermDate
{
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$True)]
        [datetime]$HireDate
    )

    if( [Boolean](Get-Random -Minimum 0 -Maximum 2) ) 
    { 
        $dt = [DateTime](Get-Random -Minimum $HireDate.Ticks -Maximum (Get-Date).Ticks)
        Write-Output $dt
    } 
}


function Get-FirstName
{
    [CmdletBinding()]
    Param(
    )
    $script:firstnames[$rand.Next(0,($firstnames.Count))].Name
}


function Get-LastName
{
    [CmdletBinding()]
    Param(
        $basepath = $PSScriptRoot #(Get-Module EmployeeFiles).ModuleBase
    )
    $script:lastnames[$rand.Next(0,($lastnames.Count))].Name
}


function Get-Employee
{
    [CmdletBinding()]
    Param(
    )

    $first = Get-FirstName
    $last  = Get-LastName
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


function Get-EmployeeFile
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$Container,

        [Parameter(Mandatory=$false)]
        [string]$Category,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputFolder = ( Join-Path ([Environment]::GetFolderPath('Desktop')) 'Samples' ),

        [Parameter(Mandatory=$false)]
        [switch]$CreateFile
    
    )

    $file = Get-Employee

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

    Write-Output $file
}

