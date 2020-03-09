﻿## isModuleAvailable
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
    #$PSModuleAutoloadingPreference = “none”

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
    $preview = new-object $typename("&Preview","Preview 0.9.6 for Template")
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
            Install-Module -Name MicrosoftTeams -RequiredVersion 0.9.6 -Repository PSGallery 
            Import-Module -Name MicrosoftTeams -RequiredVersion 0.9.6
        }
        1 {
            Write-Output $current.helpmessage
            Install-Module -Name MicrosoftTeams -RequiredVersion 1.0.3 -Repository PSGallery 
            Import-Module -Name MicrosoftTeams -RequiredVersion 1.0.3
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
        [Parameter(mandatory)][String] $domain #= "ryu365.onmicrosoft.com"
    )
    # local script 実行許可（リモートは署名付き）
    Set-ExecutionPolicy RemoteSigned

    $credential = Get-Credential $uid

    # to AzureAD  
    #Connect-MsolService  #ver.1
    Install-Module -Name "AzureAD" -Repository PSGallery
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

## 科目チーム
Function global:ruNew-ClassTeam() {
    Param (
        [parameter(mandatory)][String] $Name
    )
    # Teams Preview Module
    $tmodule = "MicrosoftTeams"
    $m = isModuleAvailable -Module $tmodule
    if ($m.Version.Major -ne 0 ) {
        Get-Module -Name $tmodule
        Write-Output "科目チーム作成には $tmodule Preview Version < 1.0 が必要です。"
        return
    }
    
    $gName = "科目_$Name"
    if (ynChoice("科目チーム「$gName」を新規作成します。") -eq 0) {
        $template = "EDU_Class"
        New-Team -DisplayName $gname -Template $template -Description $gname
    }
}

## 課程チーム
# math-course-s / g-math-course-s / math-course-t
# electro-course-s
# mecha-course-s
# material-course-s
# info-course-s
# env-course-s
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

# 入学年別プライベートチャネル
Function global:ruAdd-ChanelUser-byUid() {
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
        [Parameter(mandatory)][String]$LGroupName
    )

    $sgroup = Get-AzureADGroup -SearchString $SecurityGroupName
    $365group = Get-Team -DisplayName $LGroupName
    $365gid = $365group.GroupId
    
    $sGroupMembers = (Get-AzureADGroupMember -ObjectId $sgroup.ObjectId -All $true | select UserPrincipalName,UserType)
    $uLen = $sGroupMembers.length

    if (ynChoice("「$SecurityGroupName」の $uLen ユーザーを「$LGroupName」チームに追加します。") -eq 0) {
        foreach ($u in $sGroupMembers) {
            if ($u.UserType -eq "Member") {
                $uname = $u.UserPrincipalName
                Add-TeamUser -GroupId $365gid -User $uname
            }
        }
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
        if (ynChoice("グローバルアドレスリスト「$StaffGAL」を更新します。") -eq 0) {
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
