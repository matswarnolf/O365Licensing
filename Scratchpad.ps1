Connect-AzureAD
$users = Get-AzureADUser -all $true | Where-Object -Property AssignedPlans -ne $null
foreach ($user in $users) {
    $NewObject01 = New-Object PSObject
    $NewObject01 | Add-Member -MemberType NoteProperty -Name "Office" -Value $user.PhysicalDeliveryOfficeName
    $NewObject01 | Add-Member -MemberType NoteProperty -Name "User Principal Name" -Value $user.UserPrincipalName
    foreach ($plan in $user.AssignedPlans) {
        $number ++
        $NewObject01 | Add-Member -MemberType NoteProperty -Name "License $number" -Value $plan.Service

    }
    Remove-Variable number
    
$NewObject01 | Export-Excel -Path 'C:\Users\MatsWarnolf\OneDrive - Mats Warnolf AB\Desktop\licensereport.xlsx' -Append     
}

foreach ($user in $users){
foreach ($plan in $user.AssignedPlans){
    $NewObject01 = New-Object PSObject
    $Columndate = Get-Date -Format "MM.yyyy"
    $TableName = "A" + (New-Guid)
    $TableName = $TableName.Trimstart("-")
    $NewObject01 | Add-Member -MemberType NoteProperty -Name "Month" -Value $Columndate
    $NewObject01 | Add-Member -MemberType NoteProperty -Name "User Principal Name" -Value $user.UserPrincipalName
    $NewObject01 | Add-Member -MemberType NoteProperty -Name "Office" -Value $user.PhysicalDeliveryOfficeName
    $NewObject01 | Add-Member -MemberType NoteProperty -Name "License" -Value $plan.Service
   
    
    $NewObject01 | Export-Excel -Path 'C:\Users\MatsWarnolf\OneDrive - Mats Warnolf AB\Desktop\licensereport.xlsx' -Append  -WorksheetName "$Columndate" -TableName "$tablename" -AutoSize 
}
}
