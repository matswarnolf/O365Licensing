Connect-AzureAD
$users = Get-AzureADUser -all $true | Where-Object -Property AssignedPlans -ne $null
foreach ($user in $users) {
    $NewObject01 = New-Object PSObject
    $NewObject01 | Add-Member -MemberType NoteProperty -Name "Name" -Value $User.DisplayName
    $NewObject01 | Add-Member -MemberType NoteProperty -Name "User Principal Name" -Value $User.UserPrincipalName
    foreach ($plan in $user.AssignedPlans){
        $Number ++
        
    }
    $NewObject01 | Add-Member -MemberType NoteProperty -Name "License" -Value $User.assignedplans.Service
    $newObject01 | Export-Excel -Path 'C:\Users\MatsWarnolf\OneDrive - Mats Warnolf AB\Desktop\licensereport.xlsx' -Append     
}