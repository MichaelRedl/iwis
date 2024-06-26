# Install-Module MSOnline -Force -AllowClobber
# Install-Module AzureAD -Force -AllowClobber

# Import modules
Import-Module MSOnline
Import-Module AzureAD

# Connect to M365 and Azure AD
Connect-MsolService
Connect-AzureAD

# Get all users in the tenant using MSOnline
$allUsers = Get-MsolUser -All

# Prepare the licenses data for export
$licenses = @()
$number = 0
foreach ($user in $allUsers) {
    $number = $number + 1

        $userPrincipalName = $user.UserPrincipalName
        Write-Host $userPrincipalName -Foregroundcolor Green
        Write-Host $number -Foregroundcolor Green
        $firmenname = $null
        $abteilung = $null

        # Get user details from Azure AD
        $userDetails = Get-AzureADUser -ObjectId $user.UserPrincipalName
     #   $userDetails | Format-List * -Force
      #  Write-Host $userDetails.CompanyName -Foregroundcolor Green
      #  Write-Host $userDetails.Department -Foregroundcolor Green
        $firmenname = $userDetails.CompanyName
        $abteilung = $userDetails.Department

        $userLicenses = $user.Licenses
        foreach ($license in $userLicenses) {
            $accountSkuIdParts = $license.AccountSkuId -split ':'
            $accountSkuId = if ($accountSkuIdParts.Length -gt 1) { $accountSkuIdParts[1] } else { $accountSkuIdParts[0] }
            $licenseDetails = [PSCustomObject]@{
                Lizenz                  = $accountSkuId
                Name_Lizenzierter_User  = $userPrincipalName
                Company_Name            = $firmenname
                Department_Name         = $abteilung
            }
            $licenses += $licenseDetails
        }
}

# Sort the licenses data by the Lizenz property
$sortedLicenses = $licenses | Sort-Object -Property Lizenz

# export Licenses
$sortedLicenses | Export-Csv -Path ".\TenantUserLicenses.csv" -NoTypeInformation -Delimiter ';'
