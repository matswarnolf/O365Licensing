Connect-AzureAD
$users = Get-AzureADUser -all $true | Where-Object -Property AssignedPlans -ne $null
$NewObject02 = New-Object PSObject
$Columndate = Get-Date -Format "MM.yyyy"
foreach ($user in $users) {
    foreach ($plan in $user.AssignedPlans) {
        $NewObject01 = New-Object PSObject
       
        $NewObject01 | Add-Member -MemberType NoteProperty -Name "Month" -Value $Columndate
        $NewObject01 | Add-Member -MemberType NoteProperty -Name "User Principal Name" -Value $user.UserPrincipalName
        $NewObject01 | Add-Member -MemberType NoteProperty -Name "Office" -Value $user.PhysicalDeliveryOfficeName
        $NewObject01 | Add-Member -MemberType NoteProperty -Name "License" -Value $plan
        $NewObject02 += $NewObject01
        
    }
}
$NewObject02 | Export-Excel -Path 'C:\Users\MatsWarnolf\OneDrive - Mats Warnolf AB\Desktop\licensereport.xlsx' -Append  -WorksheetName "$Columndate" -AutoSize 