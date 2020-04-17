## isModuleAvailable
Function global:isModuleAvailable() {
    Param (
        [String] $Module
    )
    if ($m = Get-Module -ListAvailable -Name $module) {
        return $m
    } else {
        return $false
    }
}

## Yes/No Choice
Function global:ynChoice() {
    Param (
        [String] $message
    )
    #選択肢の作成
    $typename = "System.Management.Automation.Host.ChoiceDescription"
    $no  = new-object $typename("&No","実行しない")
    $yes = new-object $typename("&Yes","実行する")

    #選択肢コレクションの作成
    $assembly= $yes.getType().AssemblyQualifiedName
    $choice = new-object "System.Collections.ObjectModel.Collection``1[[$assembly]]"
    $choice.add($no)
    $choice.add($yes)

    #選択プロンプトの表示 Yes=0, No=1
    return $host.ui.PromptForChoice($message,"実行しますか？",$choice,0)
}

## セットアップ
Function global:ruInit() {
    # disable module autoloading
    $PSModuleAutoloadingPreference = “none”
    
    # ログインプロンプトをコンソールに
    Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\PowerShell\1\ShellIds" -Name ConsolePrompting -Value $true

    # AzureAD
    Install-Module -Name AzureAD -Repository PSGallery

    # Set PSGallery to Trusted
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    # Beta release Gallery
    Register-PSRepository -Name PSGalleryInt -SourceLocation https://www.poshtestgallery.com/ -InstallationPolicy Trusted

    # Install Teams Modules
    #global:TeamsChoice 
    #Install-Module -Name MicrosoftTeams -RequiredVersion 0.9.6 -Repository PSGallery -Force #-Scope CurrentUser 
    #Install-Module -Name MicrosoftTeams -Repository PSGallery -Force #-Scope CurrentUser #-RequiredVersion 1.0.3
    #Install-Module -Name MicrosoftTeams -Repository PSGalleryInt -Force #-Scope CurrentUser #-RequiredVersion 1.0.21

    #Get-InstalledModule -Name MicrosoftTeams -AllVersions
}

Function global:TeamsChoice() {
    #選択肢の作成
    $typename = "System.Management.Automation.Host.ChoiceDescription"
    $preview = new-object $typename("&Preview","Preview 0.9.10 for Template")
    $current = new-object $typename("&Current","Current version")
    $beta = new-object $typename("&Beta","Beta version for Private Channel")
    
    #選択肢コレクションの作成
    $assembly= $current.getType().AssemblyQualifiedName
    $choice = new-object "System.Collections.ObjectModel.Collection``1[[$assembly]]"
    $choice.add($preview)
    $choice.add($current)
    $choice.add($beta)

    #選択プロンプトの表示 Preview=0, Current=1, Beta=2
    $ans = $host.ui.PromptForChoice("Microsoft Teams Module for Power Shell を設定します。","どのバージョンを利用しますか？",$choice,0)

    switch ($ans) {
        0 {
            Write-Output $preview.helpmessage
            Install-Module -Name MicrosoftTeams -RequiredVersion 0.9.10 -Repository PSGallery 
            Import-Module -Name MicrosoftTeams -RequiredVersion 0.9.10
        }
        1 {
            Write-Output $current.helpmessage
            Install-Module -Name MicrosoftTeams -RequiredVersion 1.0.5 -Repository PSGallery 
            Import-Module -Name MicrosoftTeams -RequiredVersion 1.0.5
        }
        2 {
            Write-Output $beta.helpmessage
            Install-Module -Name MicrosoftTeams -RequiredVersion 1.0.21 -Repository PSGalleryInt
            Import-Module -Name MicrosoftTeams -RequiredVersion 1.0.21
        }
        default { Write-Opuput "Not matched." }
    }
    Get-InstalledModule -Name MicrosoftTeams
}

## 接続
Function global:ruConnect() {
    Param (
        [parameter(mandatory)][String] $uid,
        [parameter(mandatory)][String] $domain
    )
    # ryu365.onmicrosoft.com
    # ryuu.onmicrosoft.com / office.ryukoku.ac.jp

    # local script 実行許可（リモートは署名付き）
    Set-ExecutionPolicy RemoteSigned

    $credential = Get-Credential $uid

    # to AzureAD  
    #Connect-MsolService  #ver.1
    #Install-Module -Name "AzureAD" -Repository PSGallery
    Connect-AzureAD -Credential $credential

    # to Teams
    TeamsChoice
    Connect-MicrosoftTeams -credential $credential
    # to Skype for Buisness Online  
    # Skype Online Connector # https://www.microsoft.com/en-us/download/confirmation.aspx?id=39366
    Import-Module SkypeOnlineConnector

    $sfbsession = New-CsOnlineSession -Credential $credential –OverrideAdminDomain $domain
    Import-PSSession $sfbsession -AllowClobber

    # to Exchange online  
    $session = new-pssession -configurationName microsoft.exchange -connectionuri "https://outlook.office365.com/powershell-liveid/" -credential $credential -authentication basic -allowredirection
    import-pssession $session -disablenamechecking -AllowClobber
}

## 一般チーム
Function global:ruNew-Team() {
    Param (
        [parameter(mandatory)][String] $CharName,
        [parameter(mandatory)][String] $DisplayName,
        [parameter(mandatory)][String] $OwnerId
    )
    # Teams Current Module
    $tmodule = "MicrosoftTeams"
    $m = Get-Module -Name $tmodule
    if ($m.Version.Major -eq 0 ) {
        Write-Output $m
        Write-Output "一般チーム作成には $tmodule Version > 1.0 が必要です。"
        return
    }
    $owner = Get-AzureADUser -SearchString $OwnerId
    $ownerName = $owner.DisplayName -replace "　", "" -replace " ",""

    if (ynChoice("一般チーム「$DisplayName」を固有ID「$CharName」所有者「$ownerName」で新規作成します。") -eq 0) {
        $group = New-Team -DisplayName $CharName -MailNickName $CharName -Owner $owner.UserPrincipalName -AllowGuestCreateUpdateChannels $false -AllowGuestDeleteChannels $false -AllowCreateUpdateChannels $false -AllowDeleteChannels $false -AllowCreateUpdateRemoveTabs $false
        Set-Team -GroupId $group.GroupId -DisplayName $DisplayName
    }
}

## 科目チーム
Function global:ruNew-ClassTeam() {
    Param (
        [parameter(mandatory)][String] $ClassName,
        [parameter(mandatory)][String] $OwnerId
    )
    # Teams Preview Module
    $tmodule = "MicrosoftTeams"
    $m = Get-Module -Name $tmodule
    if ($m.Version.Major -ne 0 ) {
        Write-Output $m
        Write-Output "科目チーム作成には $tmodule Preview Version < 1.0 が必要です。"
        return
    }
    $owner = Get-AzureADUser -SearchString $OwnerId
    $ownerName = $owner.DisplayName -replace "　", "" -replace " ",""

    $gName = "科目_${ClassName}_${ownerName}"
    if (ynChoice("科目チーム「$gName」を新規作成します。") -eq 0) {
        $template = "EDU_Class"
        New-Team -DisplayName $gname -Template $template -Description $gname -Owner $owner.UserPrincipalName
    }
}

## 課程チーム
## 理工学部/理工学研究科
## math-course-s / g-math-course-s
# 4da9067a-78d8-4048-93cb-e27cf550263e
## electro-course-s
# 89e28aca-8889-4b15-a67d-82eae112aa9f
## mecha-course-s
# bf617f28-e8b1-4231-8bbf-0d73a2fcd282
## material-course-s
# 5662808f-bf17-4b29-b5f1-71519230cded
## info-course-s
# ba461383-dfb6-4eea-b92f-68829a4b7650
## env-course-s
# 5523227d-652b-40ed-8d4c-3b885b2761cc
## 先端理工教員（学生は -s ）
# Y-math-course-t
# Y-electro-course-t
# Y-mecha-course-t
# Y-material-course-t
# Y-info-course-t
# Y-env-course-t
# ExtensionAttribute でチームにメンバーを一括登録
Function global:ruAdd-TeamUser-byExtension() {
    param (
        [Parameter(mandatory)][String]$ExtString,
        [Parameter(mandatory)][String]$GroupId,
        [ValidateSet("Member","Owner")][String]$Role = "Member"
    )

    $attribute = "extension_875d2e3d99b34cab947ebf6419397ca4_extensionAttribute1"
    $users = Get-AzureADUser -All $true -Filter "startswith($attribute,'$ExtString')"
    foreach ($u in $users) {
        Write-Output $u.UserPrincipalName
    }
    $uLen = $users.length
    Write-Output "$uLen Users Found"

    $t = Get-Team -GroupId $GroupId
    $tname = $t.DisplayName

    if (ynChoice("$uLen ユーザーを「$tname」チームに追加します。") -eq 0) {
        foreach ($u in $users) {
            Add-TeamUser -GroupID $GroupId -User $u.UserPrincipalName -Role $Role
        }
        Write-Output "Done"
    }
}

# ExtensionAttributeでプライベートチャネル登録
Function global:ruAdd-ChannelUser-byExtension() {
    param (
        [Parameter(mandatory)][String]$TeamId,
        [Parameter(mandatory)][String]$ExtString,
        [Parameter(mandatory)][String]$ChannelName
    )

    $attribute = "extension_875d2e3d99b34cab947ebf6419397ca4_extensionAttribute1"
    $t_users = Get-AzureADUser -All $true -Filter "startswith($attribute,'$ExtString')"

    $uLen = $t_users.length
    Write-Output "$uLen Users Found"

    $t = Get-Team -GroupId $TeamId
    $tname = $t.DisplayName

    if (ynChoice("$uLen ユーザーを「$tname」チーム「$ChannelName」チャネルに追加します。") -eq 0) {
        foreach ($u in $t_users) {
            $uname = $u.UserPrincipalName
            #Write-Output "Add-TeamChannelUser -GroupId $TeamId -DisplayName $ChannelName -User $uname"
            Add-TeamChannelUser -GroupId $TeamId -DisplayName $ChannelName -User $uname
        }
        Write-Output "Done"
    }
}

# 入学年別プライベートチャネル
Function global:ruAdd-ChannelUser-byUid() {
    param (
        [Parameter(mandatory)][String]$TeamId,
        [Parameter(mandatory)][String]$UidString,
        [Parameter(mandatory)][String]$ExtString,
        [Parameter(mandatory)][String]$ChannelName,
        [Switch]$toOwner
    )

    $attribute = "extension_875d2e3d99b34cab947ebf6419397ca4_extensionAttribute1"
    $t_users = Get-AzureADUser -All $true -Filter "startswith($attribute,'$ExtString')"

    $c_users = @()
    #$c_users = New-Object System.Collections.ArrayList
    foreach ( $u in $t_users ) {
        if ($u.UserPrincipalName.startswith($UidString)) {
            Write-Output $u.UserPrincipalName
            $c_users += ($u)
            #$c_users.Add($u)
        }
    }
    $uLen = $c_users.length
    Write-Output "$uLen Users Found"

    $t = Get-Team -GroupId $TeamId
    $tname = $t.DisplayName

    if (ynChoice("$uLen ユーザーを「$tname」チーム「$ChannelName」チャネルに追加します。") -eq 0) {
        foreach ($u in $c_users) {
            $uname = $u.UserPrincipalName
            Write-Output "Add-TeamChannelUser -GroupId $TeamId -DisplayName $ChannelName -User $uname"
            Add-TeamChannelUser -GroupId $TeamId -DisplayName $ChannelName -User $uname
            if ($toOWner) {
                Write-Output "$uname to OWner"
                Add-TeamChannelUser -GroupId $TeamId -DisplayName $ChannelName -User $uname -Role OWner
            }
        }
        Write-Output "Done"
    }
}
    
# Securityグループメンバーを科目チームに追加
Function global:ruAdd-TeamUser-fromSecurityGroup() {
    param (
        [Parameter(mandatory)][String]$SecurityGroupName,
        [Parameter(mandatory)][String]$ClassName
    )

    $sgroup = Get-AzureADGroup -SearchString $SecurityGroupName
    $365group = Get-Team -DisplayName $ClassName
    $365gid = $365group.GroupId
    $365name = $365group.DisplayName
    
    $sGroupMembers = (Get-AzureADGroupMember -ObjectId $sgroup.ObjectId -All $true | select UserPrincipalName,UserType)
    $uLen = $sGroupMembers.length

    if (ynChoice("「$SecurityGroupName」の $uLen ユーザーを「$365name」チームに追加します。") -eq 0) {
        foreach ($u in $sGroupMembers) {
            if ($u.UserType -eq "Member") {
                $uname = $u.UserPrincipalName
                Add-TeamUser -GroupId $365gid -User $uname
            }
        }
    }
}

## ゲストアカウントを追加
# コマンド発行ユーザーからの招待になるので管理者で実行するが吉
Function global:ruAdd-GuestUser() {
    param (
        [Parameter(mandatory)][String]$DisplayName,
        [Parameter(mandatory)][String]$MailAddress,
        [ValidateSet("Faculty","Student")][String]$Role = "Faculty"
    )
    ## ryu365.onmicrosoft.com のテナントID  
    #$tenantid = 23b65fdf-a4e3-4a19-b03d-12b1d57ad76e
    #Connect-AzureAD -TenantId $tenantid
    ## AzureADPreview だと -TenantDomain が使えるらしい  
 
    $invUrl = "https://teams.microsoft.com/"
 
    if (ynChoice("ゲストユーザー $MailAddress を $DisplayName として招待します。") -eq 0) {
        New-AzureADMSInvitation -InvitedUserDisplayName $DisplayName  -InvitedUserEmailAddress  $MailAddress -InviteRedirectURL $invUrl -SendInvitationMessage $true

        $guest = Get-AzureADUser -Filter "UserType eq 'Guest' and DisplayName eq '$DisplayName'"
        Set-AzureADUser -ObjectId $guest.ObjectId -UsageLocation "JP"
        
        Write-Output "Waiting 5sec..."
        Start-Sleep -s 5

        $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
        if ($Role -eq "Faculty") {
            $license.SkuId = "94763226-9b3c-4e75-a931-5c89701abe66" # STANDARDWOFFPACK_FACULTY A1
        } else {
            $license.SkuId = "314c4481-f395-4525-be8b-2ec4bb1e9d91" # STANDARDWOFFPACK_STUDENT A1       
        }
        $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        $LicensesToAssign.AddLicenses = $License

        Set-AzureADUserLicense -ObjectId $guest.ObjectId -AssignedLicenses $LicensesToAssign
    }
}

# アドレスリストとABPを更新
Function global:ruUpdate-AddressListABP() {

    $StAddressList = "All St-Student"
    if (ynChoice("アドレスリスト「$StAddressList」を更新します。") -eq 0) {
        Remove-AddressList -Identity $StAddressList -Confirm
        New-AddressList -name $StAddressList -RecipientFilter "((RecipientType -eq 'UserMailbox' -and (Office -like 'T0*' -or Office -like 'T1*' -or Office -like 'T2*')))"
    }
    $StaffAddressList = "All Staff"
    if (ynChoice("アドレスリスト「$StaffAddressList」を更新します。") -eq 0) {
        Remove-AddressList -Identity $StaffAddressList -Confirm
        New-AddressList -name $StaffAddressList -RecipientFilter "((RecipientType -eq 'UserMailbox' -and (Office -like '0*' -or Office -like '1*' -or Office -like '2*' -or Office -like '3*' -or Office -like '4*' -or Office -like '5*' -or Office -like '6*' -or Office -like '7*' -or Office -like '8*' -or Office -like '9*')))"
    }

    if (Get-TransportConfig | Format-List AddressBookPolicyRoutingEnabled) {
        $StaffGAL = "GAL_STAFF"
        if (ynChoice("グロー
        
        バルアドレスリスト「$StaffGAL」を更新します。") -eq 0) {
            Remove-GlobalAddressList -Identity $StaffGAL -Confirm
            New-GlobalAddressList -Name $StaffGAL -RecipientFilter {(Office -like '0*' -or Office -like '1*' -or Office -like '2*' -or Office -like '3*' -or Office -like '4*' -or Office -like '5*' -or Office -like '6*' -or Office -like '7*' -or Office -like '8*' -or Office -like '9*')}
        }
        $gal = "GAL_STAFF" 
        $oab = "OfflineList" 
        $room = "RoomList" 
        $StudentABP = "Student_ABP"
        if (ynChoice("アドレスブックポリシー「$StudentABP」を更新します。") -eq 0) {
            Remove-AddressBookPolicy -Identity $StudentABP
            New-AddressBookPolicy -Name $StudentABP -Addresslists $StaffAddressList -OfflineAddressBook $oab -GlobalAddressList $gal -RoomList $room
        }

        if (ynChoice("ユーザーのアドレスブックポリシーとGALを更新します。") -eq 0) {
            Write-Output "Set ABP" 
            $stbox = Get-Mailbox -ResultSize unlimited -Filter {Office -like 'T0*' -or Office -like 'T1*' -or Office -like 'T2*'}
            $stbox | foreach { Set-Mailbox -Identity $_.Identity -AddressBookPolicy $StudentABP }
            Write-Output "Update User GAL" 
            foreach ($u in @("0*","1*","2*","3*","4*","5*","6*","7*","8*","9*")) {
                Write-Output "Updating User $u"
                Get-Recipient -ResultSize unlimited | ? { $_.Office -like $u } | foreach {Update-Recipient -Identity $_.Identity}
            }
        }

    } else {
        Write-Output "AddressBookPolicyRoutingEnabled : False"
        Write-Output "Do Notihg"
    }
} 
