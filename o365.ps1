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
Function global:ruSetup() {
    param (
        [switch]$TeamsPreview
    )
    # local script 実行許可（リモートは署名付き）
    Set-ExecutionPolicy RemoteSigned

    # PSGallery
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    #Register-PSRepository -Name PSGalleryInt -SourceLocation https://www.poshtestgallery.com/ -InstallationPolicy Trusted

    # Basic Modules
    $modules = "AzureAD"#, "MSOnline" #ver.1
    foreach ( $module in $modules) {
        if (-Not (isModuleAvailable -Module $module) ) {
            Install-Module -Name $module
            Get-Module -Name $module
        }
    }

    # Teams Module
    $tmodule = "MicrosoftTeams"
    if ($m = isModuleAvailable -Module $tmodule) {
        if (($TeamsPreview) -And ($m.Version.Major -ne 0 )) {
            # reinstall preview version
            #Disconnect-MicrosoftTeams
            Uninstall-Module -Name MicrosoftTeams
            Install-Module -Name MicrosoftTeams -RequiredVersion 0.9.6 -Force
        } elseif (-Not ($TeamsPreview) -And ($m.Version.Major -eq 0)) {
            # reinstall preview version
            #Disconnect-MicrosoftTeams
            Uninstall-Module -Name MicrosoftTeams
            Install-Module -Name MicrosoftTeam -Force
        }
    } else {
        if ($TeamsPreview) {
            Install-Module -Name MicrosoftTeams -RequiredVersion 0.9.6 -Force
        } else {
            Install-Module -Name MicrosoftTeams
        }
    }
    Get-Module -Name $tmodule

    # Skype Online Connector
    #Import-Module "C:\\Program Files\\Common Files\\Skype for Business Online\\Modules\\SkypeOnlineConnector\\SkypeOnlineConnector.psd1" 
    Import-Module SkypeOnlineConnector
}

## 接続
Function global:ruConnect() {
    Param (
        [parameter(mandatory)][String] $uid,
        [Parameter(mandatory)][String] $domain #= "ryu365.onmicrosoft.com"
    )
    #$uid = "a00007@mail.ryukoku.ac.jp"
    #$domain = "ryu365.onmicrosoft.com"

    $credential = Get-Credential $uid

    # to AzureAD  
    #Connect-MsolService  #ver.1
    Connect-AzureAD -Credential $credential

    # to Teams  
    Connect-MicrosoftTeams -credential $credential
    # to Skype for Buisness Online  
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
