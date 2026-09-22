Connect-AzAccount -Tenant '<tenantId>'
Set-AzContext -Subscription '<subscriptionId>'

$subscriptionId = '<subscriptionId>'
$scope = "/subscriptions/$subscriptionId"



$spObjectId = (Get-AzADServicePrincipal -ApplicationId '<appClientId>').Id
$spObjectId   # confirm it returned a GUID



$role = @{
    Name             = 'PIM Assignment Schedule Reader'
    IsCustom         = $true
    Description      = 'Read PIM for Azure resources role assignment schedule instances.'
    Actions          = @(
        'Microsoft.Authorization/roleAssignmentScheduleInstances/read',
        'Microsoft.Authorization/roleDefinitions/read'
    )
    NotActions       = @()
    DataActions      = @()
    NotDataActions   = @()
    AssignableScopes = @($scope)
}
$role | ConvertTo-Json | Set-Content -Path .\pim-schedule-reader.json
New-AzRoleDefinition -InputFile .\pim-schedule-reader.json



New-AzRoleAssignment -ObjectId $spObjectId `
    -RoleDefinitionName 'PIM Assignment Schedule Reader' `
    -Scope $scope
Repeat for the second app's $spObjectId if you have one.



Get-AzRoleAssignment -ObjectId $spObjectId -Scope $scope |
    Select-Object RoleDefinitionName, Scope
