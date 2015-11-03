<#
.Synopsis
   This function finds the manager of AD groups and that users business manager
.DESCRIPTION
   This function looks at a specific OU/Sub-OU's and grabs the following information for notification processing:
   GroupName
   GroupManager
   The actual manager of the user that is managing the group
   The actual manager's email of the user that is managing the group
.PARAMETER SearchBase
    This parameter set's the scope of the AD Group search.  Please provide a OU structured string to begin the search.
.EXAMPLE
   Get-ADGroupMemberManager -SearchBase 'OU=Managed,OU=Groups,OU=Division Resources,DC=some,DC=domain,DC=com'
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   This function outputs an object that contains:
        ADGroupObject = The AD Group Object that has all the information setup correctly and identifiable
        NoEmailObject = This object contains all of the groups that do not have an email listed for the group manager manager
        NoManagerObject = This object contains all the groups that do not have a manager listed for the group itself.

#>
function Get-ADGroupMemberManager
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("OUPath")] 
        $SearchBase,

        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("OutPath")] 
        $FolderPath = ([Environment]::GetFolderPath("Desktop")),

        [Parameter()]
        [switch]$WhatIf
    )

    $results = Find-ADGroupMemberManager -SearchBase $SearchBase -FolderPath $FolderPath

    if ($results.ADGroupObject)
    {
        $results.ADGroupObject | Select -Property * | Export-Csv $FolderPath\ADGroupObject.csv
    }
    if ($results.NOEmailObject)
    {
        $results.NoEmailObject | Select -Property * | Export-Csv $FolderPath\NoEmailObject.csv
    }
    if ($results.NoManagerObject)
    {
        $results.NoManagerObject | Select -Property * | Export-Csv $FolderPath\NoManagerObject.csv
    }

    if ($WhatIf)
    {
        Send-ADGroupMemberManagerNotification -inputobject $results -WhatIf -FolderPath $FolderPath
    }
    else
    {
        Send-ADGroupMemberManagerNotification -inputobject $results
    }
}
