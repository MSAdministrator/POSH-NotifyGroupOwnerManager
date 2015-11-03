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
   Get-ADGroupMemberManager -SearchBase 'OU=,OU=Groups,OU=Division Resources,DC=some,DC=domain,DC=com'
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   This function outputs an object that contains:
        ADGroupObject = The AD Group Object that has all the information setup correctly and identifiable
        NoEmailObject = This object contains all of the groups that do not have an email listed for the group manager manager
        NoManagerObject = This object contains all the groups that do not have a manager listed for the group itself.

#>
function Find-ADGroupMemberManager
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
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
        $FolderPath = ([Environment]::GetFolderPath("Desktop"))
    )

    Begin
    {
        $ADGroupObject = @()
        $NoEmailObject = @()
        $NoManagerObject = @()
        $adgroupcount = @()
        $i = 0

        $adgroup = Get-ADGroup -Filter * -SearchBase $SearchBase -Properties *

        foreach ($group in $adgroup)
        {
            $adgroupcount += $group
        }
    }
    Process
    {
        foreach ($item in $adgroup)
        {
            $UsersManager = ''
            $ManagedByManagerEmail = ''

            Write-Progress -Activity "Gathering information from $($($adgroupcount).count) AD Groups" -Status "Processing group $($item.Name)" -PercentComplete ($i/$($($adgroupcount).count)*100)

            if (![string]::IsNullOrWhiteSpace($item.ManagedBy))
            {
                Write-Verbose "Group $($item.name) is managed by $($item.ManagedBy)"
                try
                {
                    #additional Properties that could be selected: objectGUID, displayName, office, division, department, employeeNumber, employeeID, mobilePhone, officePhone, ipphone, title, givenName, surname,mail
                    $UsersManager = Get-ADUser -Identity $($item.ManagedBy) -Properties * | select sAMAccountName, @{Name='Manager';Expression={(Get-ADUser $_.Manager).sAMAccountName}}
                }
                catch
                {
                    $msg=('An error occurred that could not be resolved: {0}' –f $_.Exception.Message)
                    Write-Warning $msg
                    #Write the exception to a log file
                    $_.Exception | Select-Object * | Out-file "$($FolderPath)\errors.txt" –append
                    #Export the error to XML for later diagnosis
                    $_ | Export-Clixml "$($FolderPath)\UnknownException.xml"
                }
                
                
                Write-Verbose "User: $($UsersManager.sAMAccountName) manager is $($UsersManager.manager)"
                try
                {
                    $ManagedByManagerEmail = "$((Get-ADUser -Identity $UsersManager.manager -Properties * | select mail).mail)"
                }
                catch
                {
                    $msg=('An error occurred that could not be resolved: {0}' –f $_.Exception.Message)
                    Write-Warning $msg
                    #Write the exception to a log file
                    $_.Exception | Select-Object * | Out-file "$($FolderPath)\errors.txt" –append
                    #Export the error to XML for later diagnosis
                    $_ | Export-Clixml "$($FolderPath)\UnknownException.xml"
                }

                if (![string]::IsNullOrWhiteSpace($ManagedByManagerEmail)){
            
                    $props = [ordered]@{
                        GroupName = $($item.Name)
                        GroupDescription = $($item.Description)
                        GroupManagedBy = $((Get-ADUSer $item.ManagedBy).Name)
                        ManagedByManager = $($UsersManager.manager)
                        ManagedByManagerEmail = $ManagedByManagerEmail
                    }

                    $tempADGroupObject = New-Object -TypeName PSObject -Property $props
                    $ADGroupObject += $tempADGroupObject
                }
                else
                {
                    $noemailprops = [ordered]@{
                        GroupName = $($item.Name)
                        GroupDescription = $($item.Description)
                        GroupManagedBy = $((Get-ADUSer $item.ManagedBy).Name)
                        ManagedByManager = $($UsersManager.manager)
                    }

                    $tempNoEmailObject = New-Object -TypeName PSObject -Property $noemailprops
                    $NoEmailObject += $tempNoEmailObject
                }
            }
            else
            {
                $nomanagerprops = [ordered]@{
                        GroupName = $($item.Name)
                        GroupDescription = $($item.Description)
                        NoManagedByUser = 'This group is not managed by a user'
                    }

                    $tempNoManagerObject = New-Object -TypeName PSObject -Property $nomanagerprops
                    $NoManagerObject += $tempNoManagerObject
            }
        $i++
        }
    }
    End
    {
        $ADObjectProps = @{
            ADGroupObject = $ADGroupObject
            NoEmailObject = $NoEmailObject
            NoManagerObject = $NoManagerObject
        }

        $results = New-Object -TypeName PSObject -Property $ADObjectProps

        return $results
    }
}
