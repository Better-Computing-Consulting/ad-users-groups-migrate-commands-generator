Get-ADUser -Filter {enabled -eq $true} -properties * | ? {$_.mail -like "*source.org"} | `
    select Name, GivenName, SamAccountName, DisplayName, DistinguishedName, UserPrincipalName, `
    Surname, OtherName, Initials, City, State, Company, Country, Office, Department, Description, Title, EmailAddress | `
    Export-Csv -Path C:\temp\users.txt -Delimiter "`t" -NoTypeInformation

$users = Get-ADUser -Filter {enabled -eq $true} -Properties name, mail | ? {$_.mail -like "*source.org"}
foreach ($user in $users){
    $groupsdns = get-aduser $user.DistinguishedName -Properties MemberOf | select -ExpandProperty MemberOf
    $groups = @()
    foreach ($group in $groupsdns){ $groups += Get-ADGroup $group }
    foreach ($g in $groups){ 
        [string]::Format("{0}`t{1}`t{2}`t{3}`t{4}",$user, $g.Name, $g.DistinguishedName, $g.GroupCategory, $g.GroupScope) | `
        Out-File c:\temp\groups.txt -Append}
}