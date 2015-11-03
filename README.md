# POSH-NotifyGroupOwnerManager
This module is for notifying the manager of the person that owns a group to verify that the correct ownership is setup.

This function looks at a specific OU/Sub-OU's and grabs the following information for notification processing:
* GroupName
* GroupManager
* The actual manager of the user that is managing the group
* The actual manager's email of the user that is managing the group

# EXAMPLE
* Get-ADGroupMemberManager -SearchBase 'OU=Managed,OU=Groups,OU=Division Resources,DC=some,DC=domain,DC=com'

# OUTPUTS
This function outputs an object that contains:
   
* ADGroupObject = The AD Group Object that has all the information setup correctly and identifiable
* NoEmailObject = This object contains all of the groups that do not have an email listed for the group manager manager
* NoManagerObject = This object contains all the groups that do not have a manager listed for the group itself.
