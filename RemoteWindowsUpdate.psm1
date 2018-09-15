# 概要
# 
# リモートホストで Windows Update を実行する
# 
# 
# 詳細
# 
# RemoteWindowsUpdate.psm1 は、リモートホストで Windows Update を実行するために 
# Invoke-WindowsUpdate コマンドレットを利用可能にする。
# 
# Invoke-WindowsUpdate コマンドレットは、あるドメインに所属するホストから同じドメイン
# に所属するリモートホストに対して実行されることを想定している。そのため、認証情報を受け
# 取る Credential パラメータを持たないが、Windows Update を実行するために管理者
# （Domain Admins 相当）の権限を必要とする。
# 
# Invoke-WindowsUpdate は、Windows Update を実行するスクリプトを対象のリモートホスト
# に PSScheduledJob として登録・実行する。
# 
# スクリプトが完了すると、PSScheduledJob は削除され、対象のリモートホストに痕跡は残らない。
# 
# AutoReboot スイッチが有効な場合、Windows Update が完了すると、再起動要否が確認され、
# 必要であれば対象のリモートホストは再起動される。
# 
# 最後に、利用可能な Update の有無が確認される。利用可能な Update が無い場合、リモート
# ホストに適用されている Update の一覧がローカルホストのカレントディレクトリに出力される。
# ジョブとして実行されている場合は、完了ステータスが、Completed になる。
# 利用可能な Update が残っている場合、ローカルホストのカレントディレクトリにその一覧が
# 出力され、アップデート未完了の例外が発生する。従って、ジョブとして実行されている場合は、
# 完了ステータスが Failed になる。
# 
# 完了ステータスが Failed になった場合は、単純に Invoke-WindowsUpdate を再実行するだけ
# で、たいてい解決する。問題がある場合はトラブルシューティングの項目を参照のこと。
# 
# 
# 使い方
# 
# Import-Module -Verbose -Force RemoteWindowsUpdate.psm1
# Invoke-WindowsUpdate -AsJob -AutoReboot -ComputerName {対象ホスト名}
# 
# 
# 利用要件
# 
# Windows Server 2012R2 以上
# もしくは PowerShell 3.0 以上で WinRM が利用可能であること
# 
# 
# トラブルシューティング
# 
# Invoke-WindowsUpdate を実行すると、すぐに異常終了してしまう場合、同名のタスクが
# 実行中である可能性がある。先行しているタスクが終了するのを待ってから再実行する。
# タスクスケジューラで \Microsoft\Windows\PowerShell\ScheduledJob\ を確認する。
# 

# PowerShell のバックグラウンドジョブは、元のスクリプトとは別プロセスになるため
# バックグラウンドジョブの内部で必要な関数は、全て $InitScript にまとめている。
$InitScript = {
    # Windows Update(Microsoft Update) を実行する
    Function Install-WindowsUpdate {
        Param(
            [Parameter(Position=1,Mandatory=$False)]
            [switch]$ListOnly
        )
        $ErrorActionPreference = "Stop"

        # アップデートを検索する
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $searcher = $updateSession.CreateUpdateSearcher()

        # 参照するサーバはグループポリシーで設定済みの場合は、デフォルトを選択する。
        $searcher.ServerSelection = 0

        # 未インストールの項目を全て検索する。
        $searchResult = $searcher.search("IsInstalled=0")

        if ($ListOnly) {
            # ListOnly スイッチが True
            return $searchResult.Updates
        }
    
        if ($searchResult.Updates.Count -eq 0) {
            # Update なし
            return
        }

        # EULA を承諾してダウンロードリストに追加する
        # WSUSで承認済みのため承諾は自動でOK
        $updatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
        $searchResult.Updates | ForEach-Object {
            $update = $_
            if ($update.InstallationBehavior.CanRequestUserInput) {
                # 手作業が必要
                return
            }
            if (-not $_.EulaAccepted) {
                # EULA を承諾する
                $_.AcceptEula()
            }
            $updatesToDownload.add($update) | Out-Null
        }

        if ($updatesToDownload.Count -eq 0) {
            # 手作業でEULAの承諾が必要
            throw "Could not accept the EULA"
        }

        # ダウンロードを開始
        $downloader = $updateSession.CreateUpdateDownloader()
        $downloader.Updates = $updatesToDownload
        $downloader.Download() | Out-Null

        # ダウンロードが完了した Update をインストールリストに追加する
        $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
        $searchResult.Updates | Where-Object {
            $_.IsDownloaded
        } | ForEach-Object {
            $updatesToInstall.add($_) | Out-Null
        }

        if ($updatesToInstall.Count -eq 0) {
            # ダウンロード失敗
            throw "Update Download Faild"
        }

        # "インストール開始"
        $installer = $updateSession.CreateUpdateInstaller()
        $installer.Updates = $updatesToInstall
        $installationResult = $installer.Install()

        if ($installationResult.ResultCode -ne 2) {
            # "インストール失敗"
            throw "Update Install Faild"
        }
    }
    Function Install-RemoteWindowsUpdate {
        Param(
            [Parameter(Position=1,Mandatory=$True)]
            [string]$Script
        )
        $ErrorActionPreference = "Stop"

        $ScriptBlock = Invoke-Expression $Script

        $TaskPath = "\Microsoft\Windows\PowerShell\ScheduledJobs\"
        $DefinitionName = "RemoteWindowsUpdate"

        # 既に登録されているジョブがあるか確認する
        $Job = Get-ScheduledJob -Name $DefinitionName -ErrorAction "SilentlyContinue"
        if ($Job) {
            # タスクスケジューラ上で、ジョブが稼動している場合はエラー
            $task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $DefinitionName
            if ($task.state -eq "Running") {
                throw "Windows Updating..."
            }
            # 既存のジョブを削除
            Unregister-ScheduledJob -Name $DefinitionName
        }

        # 最上位権限で実行するオプション
        $opt = New-ScheduledJobOption -RunElevated
        
        # ジョブをタスクに登録し、タスクとして実行する。
        # これはリモートセッションの権限で Windows Update を実行すると、
        # WSUS からダウンロードしようとする際にアクセスエラーが発生するため
        # タスクスケジューラを使って間接的に実行する。
        Register-ScheduledJob -ScriptBlock $ScriptBlock -Name $DefinitionName -ScheduledJobOption $opt -RunNow | Out-Null

        # ジョブ(タスク)の終了を待つ
        $Task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $DefinitionName
        while ($Task.State -eq "Running") {
            Start-Sleep -Seconds 5
            $Task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $DefinitionName
        }

        # ジョブを削除する
        Unregister-ScheduledJob -Name $DefinitionName

        # 再起動要否を返す
        Test-Path "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
    }
}

$RemoteWindowsUpdate = {
    Param(
        [Parameter(Position=1,Mandatory=$True)]
        [string]$ComputerName,

        [Parameter(Position=2,Mandatory=$False)]
        [switch]$AutoReboot,

        [Parameter(Position=3,Mandatory=$False)]
        [switch]$ListOnly
    )

    $ErrorActionPreference = "Stop"

    # スクリプトをリモートで実行
    # 戻り値として再起動要否を受け取る
    $RebootRequired = Invoke-Command -ComputerName $ComputerName -ScriptBlock ${Function:Install-RemoteWindowsUpdate} -ArgumentList ("{" + ${Function:Install-WindowsUpdate}.ToString() + "}")

    # 必要があれば再起動する
    if ($RebootRequired -and $AutoReboot) {
        Restart-Computer $ComputerName -Wait -Force
    }

    # 追加のアップデートを検索
    $AvailableUpdates = Invoke-Command -ComputerName $ComputerName -ScriptBlock ${Function:Install-WindowsUpdate} -ArgumentList $True

    $FilePath = "RemoteWindowsUpdate_{0}_{1}.txt" -f $ComputerName,(Get-Date -f "yyyyMMdd")
    if ($AvailableUpdates) {
        $updates = $AvailableUpdates | Select-Object @{L="KB";E={$_.KBArticleIds -join ","}},Title,LastDeploymentChangeTime
        $updates | ConvertTo-Csv -NTI | Out-File -Encoding Default -FilePath $FilePath -Force
        throw "$ComputerName Update Uncomplete"
    } else {
        $hotfix = Get-HotFix -ComputerName $ComputerName | Sort-Object InstalledOn
        $hotfix | ConvertTo-Csv -NTI | Out-File -Encoding Default -FilePath $FilePath -Force
        Write-OutPut "$ComputerName Update Complete"
    }
}

Function Invoke-WindowsUpdate {
    Param(
        [Parameter(Position=1,Mandatory=$True)]
        [string]$ComputerName,

        [Parameter(Position=2,Mandatory=$False)]
        [switch]$AutoReboot,

        [Parameter(Position=3,Mandatory=$False)]
        [switch]$AsJob,

        [Parameter(Position=4,Mandatory=$False)]
        [switch]$ListOnly
    )
    if ($AsJob) {
        Start-Job -Name $ComputerName -InitializationScript $InitScript -ScriptBlock $RemoteWindowsUpdate -ArgumentList $ComputerName,$AutoReboot,$ListOnly | Out-Null
    } else {
        Start-Job -Name $ComputerName -InitializationScript $InitScript -ScriptBlock $RemoteWindowsUpdate -ArgumentList $ComputerName,$AutoReboot,$ListOnly | Wait-Job | Receive-Job -AutoRemoveJob -Wait
    }
}

Export-ModuleMember -Function "Invoke-WindowsUpdate"
