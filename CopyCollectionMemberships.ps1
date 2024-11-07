
$SiteCode = Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation | Select-Object -ExpandProperty SiteCode
$ProviderMachineName = (Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation).PSComputerName # SMS Provider machine name

$initParams = @{}

# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams

# Prompt for source and destination collection names
$sourceCollectionName = Read-Host "Enter the name of the source collection"
$destinationCollectionName = Read-Host "Enter the name of the destination collection"

# Import the SCCM module
Import-Module "$($Env:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"

# Get the collection IDs based on the names
$sourceCollection = Get-CMDeviceCollection -Name $sourceCollectionName
$destinationCollection = Get-CMDeviceCollection -Name $destinationCollectionName

if ($sourceCollection -and $destinationCollection) {
    # Copy included collections
    $includedCollections = Get-CMDeviceCollectionIncludeMembershipRule -CollectionId $sourceCollection.CollectionID
    foreach ($collection in $includedCollections) {
        Add-CMDeviceCollectionIncludeMembershipRule -CollectionId $destinationCollection.CollectionID -IncludeCollectionId $collection.IncludeCollectionID
    }

    # Copy excluded collections
    $ExcludedCollections = Get-CMDeviceCollectionExcludeMembershipRule -CollectionId $sourceCollection.CollectionID
    foreach ($collection in $ExcludedCollections) {
        Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $destinationCollection.CollectionID -ExcludeCollectionId $collection.ExcludeCollectionID
    }

    # Get direct members of the source collection
    $sourceMembers = Get-CMCollectionMember -CollectionId $sourceCollection.CollectionID
    # Add members to the destination collection
    foreach ($member in $sourceMembers) {
        if($member.IsDirect){
            Add-CMDeviceCollectionDirectMembershipRule -CollectionId $destinationCollection.CollectionID -ResourceId $member.ResourceID
        }
    }

    #Get querys of the sources collection
    $queryRules = Get-CMDeviceCollectionQueryMembershipRule -CollectionId $sourceCollection.CollectionID 
    foreach ($rule in $queryRules) 
    { 
        Add-CMDeviceCollectionQueryMembershipRule -CollectionId $destinationCollection.CollectionID -QueryExpression $rule.QueryExpression -RuleName $rule.RuleName 
    }

    Write-Output "Members copied from collection '$sourceCollectionName' to '$destinationCollectionName'."
} 
else {
    Write-Output "Could not find one or both of the specified collections."
}
