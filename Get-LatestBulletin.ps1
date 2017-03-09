Function Find-Bulletin
{
    [cmdletbinding()]
    Param
    (
        [Object]$ExcelObject,
        [string]$ProductFilter,
        [string]$BulletinId
    )
    #Write-Verbose "Searching supersedes for $BulletinId."
    $oExcel | Where-Object{($_."Affected Product" -eq $ProductFilter) -and ($_.'Bulletin Id' -eq "$BulletinId")}
}

function Find-LatestBulletin
{
    [cmdletbinding()]
    Param
    (
        [Object]$ExcelObject,
        [string]$ProductFilter,
        [string]$BulletinId
    )
    Write-Verbose "Searching supersedes for $BulletinId."
    $oExcel | Where-Object{($_."Affected Product" -eq $ProductFilter) -and ($_.Supersedes -like "$BulletinId*")}
}


<#
.SYNOPSIS
   Searches "BulletinSearch.xlsx" for lastest bulletin.
.DESCRIPTION
   Microsoft publish an Excel file containing all security bulletins, past and present. As bulletins are superseded, it becomes an arduous task to find
   the latest version. This script will search through Microsoft's security bulletin data (Excel file) to find the latest version of the given bulletin.
.NOTES
    Download BulletinSearch.xlsx from https://www.microsoft.com/en-us/download/details.aspx?id=36982
    Requires ImportExcel module. Search https://www.powershellgallery.com/packages for "ImportExcel" by Doug Finke.
    Author: Adrian Walker
    Last update: 2017-03-09
.PARAMETER ExcelFile
    Full path and name to the Excel file, downloaded from https://www.microsoft.com/en-us/download/details.aspx?id=36982.
.PARAMETER BulletinId
    The bulletinId you're interested in. e.g MS13-049.
.PARAMETER ProductFilter
    Taken from the "Affected Product" column in the aforementioned Excel file. Must match exactly. 
.EXAMPLE
   ./Get-LatestBulletin -ExcelFile C:\Temp\BulletinSearch.xlsx -BulletinId MS13-049 -ProductFilter "Windows Server 2008 R2 for x64-based Systems Service Pack 1"
   Returns the latest bulletin, including all those that have been superseded since, that are applicable to "Windows Server 2008 R2 for x64-based Systems Service Pack 1"
.EXAMPLE
   ./Get-LatestBulletin -ExcelFile C:\Temp\BulletinSearch.xlsx -BulletinId MS08-069 -ProductFilter "Windows 7 for x64-based Systems Service Pack 1"
   Returns the latest bulletin, including all those that have been superseded since, that are applicable to "Windows 7 for x64-based Systems Service Pack 1".
#>
Function Get-LatestBulletin
{
    [cmdletbinding()]
    Param
    (
        [parameter(Mandatory,Position=0)]
        [string]$ExcelFile,
        [parameter(Mandatory,Position=1)]
        [string]$BulletinId,
        [parameter(Mandatory,Position=2)]
        [string]$ProductFilter
    )

   
    if(-not($oExcel))
    {
        # ImportExcel Module required: https://www.powershellgallery.com/packages/ImportExcel/2.2.9
        $oExcel = Import-Excel -Path $ExcelFile -WorkSheetname "BulletinSearch"
    }
    else
    {
        Write-Verbose "Excel sheet already loaded."
    }

    $oData = @()
    $BulletinIdSearch = $BulletinId
    Write-Verbose "Searching for $BulletinIdSearch"
    $Bulletin = Find-LatestBulletin -ExcelObject $oExcel -BulletinId $BulletinIdSearch -ProductFilter $ProductFilter
    
    #Add first search result to object:
    if($Bulletin)
    {
        $oData += [pscustomobject]@{BulletinId=$Bulletin.Supersedes.split("[")[0];SupersededBy=$Bulletin.'Bulletin Id';BulletinKB = $Bulletin.'Bulletin KB'}

        Do
        {
            $Bulletin = Find-LatestBulletin -ExcelObject $oExcel -BulletinId $Bulletin.'Bulletin Id' -ProductFilter $ProductFilter
            if($Bulletin.Supersedes)
            {
                $oData += [pscustomobject]@{
                    #These are intentionally the wrong way round, so it reads better.
                    BulletinId = $Bulletin.Supersedes.split("[")[0]
                    SupersededBy = $Bulletin.'Bulletin Id'
                    BulletinKB = $Bulletin.'Bulletin KB' 
                }
            }
    
        }While($Bulletin.Supersedes)
        
        #Add final result:
        $LastRecord = $oData | Select -Last 1
        $oLastRecord = Find-Bulletin -BulletinId $LastRecord.SupersededBy -ExcelObject $oExcel -ProductFilter $ProductFilter
        $oData += [pscustomobject]@{BulletinId=$oLastRecord.'Bulletin Id';BulletinKB=$oLastRecord.'Bulletin KB';SupersededBy="Latest"}
    }
    else
    {
        $IsLatest = Find-Bulletin -BulletinId $BulletinIdSearch -ExcelObject $oExcel  -ProductFilter $ProductFilter
        $oData += [pscustomobject]@{BulletinId=$IsLatest.'Bulletin Id';BulletinKB=$IsLatest.'Bulletin KB';SupersededBy="Latest"}
    }
    $oData
    
}



