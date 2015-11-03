<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
function Send-ADGroupMemberManagerNotification
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
        [object]$inputobject,
        [Parameter()]
        [switch]$WhatIf,
        
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("OutPath")] 
        $FolderPath = ([Environment]::GetFolderPath("Desktop"))
    )

    Clear-Host
    try
    {
        Add-Type -assembly "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
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


        $ADGroups = @()
        
        [array]$managerEmail = @()
        $managerEmail = $inputobject.ADGroupObject.ManagedByManagerEmail | select -Unique
        $managerEmail.count

        [array]$groupManagedBy = @()
        $groupManagedBy = $inputobject.ADGroupObject.GroupManagedBy
   
        for ($i = 0; $i -lt ($managerEmail).count; $i++)
        {
            for ($g = 0; $g -lt ($inputobject.ADGroupObject.GroupManagedBy).count; $g++)
            {               
                foreach ($item in $inputobject.ADGroupObject)
                {
                    if ($($item.ManagedByManagerEmail) -eq $managerEmail[$i])
                    {   
                        if ($($item.GroupManagedBy) -eq $groupManagedBy[$g])
                        {  
                            $ADGroupName += $($item.GroupName)
                            $ADGroupManagedBy += $($item.GroupManagedBy)

                            $ADGroups += "<tr><td>$($item.GroupName)</td><td>$($item.Description)</td><td>$($item.GroupManagedBy)</></tr>"
                        }
                    }
                }
            }
         




        $html = @" 
<!DOCTYPE html>
<html>
<head>
<title>Electronic Resource Review - Group Managers</title>



<style type="text/css">
    @media only screen and (max-width: 480px){
        .emailButton{
            max-width:600px !important;
            width:100% !important;
        }

        .emailButton a{
            display:block !important;
            font-size:18px !important;
        }
    }

    .image-and-text{
    clear:both;
}
.image{
    float:right;
}
.text{
    float:left;
}

.full {
     width:100%;
     height:auto;
     }

</style>




</head>
<body style="color: #000000; font-family: Arial, sans-serif; font-size: 12px; line-height: 20px;">

<table border="0" cellpadding="0" cellspacing="0" width="100%">
<div id="header" >
        <div style="float:left; margin-top:20px" >
            <img src="https://raw.githubusercontent.com/MSAdministrator/QualysGuard-V1-API---PowerShell/master/images/doit_logo.jpg" height="70" alt="logo" align="left" />
        </div>
        <div style="float:right; margin-top:20px" >
            <img src="https://raw.githubusercontent.com/MSAdministrator/QualysGuard-V1-API---PowerShell/master/images/mu_logo.png" height="70" alt="logo" align="right" />
        </div>
    </div>
</table>
<table border="0" cellpadding="0" cellspacing="0" class="emailButton" style="-webkit-border-radius: 3px; -moz-border-radius: 3px; border-radius: 5px; background-color:#505050; border:1px solid #353535;" width="100%" arcsize="13%">
    <tr>
        <td align="center" valign="middle" style="color:#FFFFFF; font-family:Helvetica, Arial, sans-serif; font-size:16px; font-weight:bold; letter-spacing:-.5px; line-height:150%; padding-top:15px; padding-right:30px; padding-bottom:15px; padding-left:30px;">
            <a style="color:#FFFFFF; text-decoration:none;">Electronic Resource Review - Group Managers</a>
        </td>
    </tr>
</table>
</div>
<p>Date: $(Get-Date) </p>

<table border="0" cellpadding="0" cellspacing="0" width="100%" style="color:#00000; font-family:Helvetica, Arial, sans-serif; font-size:12px; line-height:150%; padding-top:5px; padding-right:5px; padding-bottom:5px; padding-left:5px;">
    <tr>
        <td style="padding:5px;">
            <p><b>Hello $($managerEmail[$i])
            <p><b>The following information lists groups which control access to electronic resources at the University of Missouri. You are receiving this message since the groups are managed (owned) by individuals who report to you, per the University's Global Address List (GAL).</b></p>
            <p><b>Please review the list and reply back with group manager changes if they are needed.</b></p>    
               

	<table align="middle" style="color:#00000; font-family:Helvetica, Arial, sans-serif; font-size:12px; line-height:150%; padding-top:5px; padding-right:5px; padding-bottom:5px; padding-left:5px;">
		<tr>
            <th bgcolor="#A0A0A0"><center><b>Active Directory Group</b>:</th>
            <th bgcolor="#A0A0A0"><center><b>Description</b>:</th>
            <th bgcolor="#A0A0A0"><center><b>User with Manager rights over Group</b>:</th>
        </tr>
        $($ADGroups)
	</table>
</table>
</body>
</html>
"@

        if ($WhatIf)
        {
            write-host "ManagerEmail: $($managerEmail[$i])"
            Write-Host "ADGroups: $($ADGroups) " `n
            $ADGroups = @()
        }
        else
        {
            $Outlook = New-Object -ComObject Outlook.Application

            $Mail = $Outlook.CreateItem(0)
            #$Mail.To = "$($managerEmail[$i])"
            $Mail.To = 'rickardj@missouri.edu'
            #$Mail.Sentonbehalfofname = "abuse@missouri.edu"
            $Mail.Subject = "[DRAFT]AD Group Owner Verification"
    
            $Mail.HTMLBody =$html
            $Mail.Send()
 
            $ADGroups = @()
        }
    }
    Start-Sleep -Seconds 5
    Get-Process -Name OUTLOOK | Stop-Process    
}
