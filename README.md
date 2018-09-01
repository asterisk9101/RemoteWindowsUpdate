# RemoteWindowsUpdate

リモートホストで Windows Update を実行する

# 詳細

RemoteWindowsUpdate.psm1 は、リモートホストで Windows Update を実行するために 
Invoke-WindowsUpdate コマンドレットを利用可能にする。

Invoke-WindowsUpdate コマンドレットは、あるドメインに所属するホストから同じドメイン
に所属するリモートホストに対して実行されることを想定している。そのため、認証情報を受け
取る Credential パラメータを持たないが、Windows Update を実行するために管理者
（Domain Admins 相当）の権限を必要とする。

Invoke-WindowsUpdate は、Windows Update を実行するスクリプトを対象のリモートホスト
に設置する。特に変更しなければリモートホストに C:\Invoke-RemoteTask.ps1 が作成される。

続いて、対象のリモートホストにスクリプトを実行するタスクをタスクスケジューラのルートに
作成し、実行する。スクリプトを直接実行しないのは Windows Update がリモート実行できな
いための回避策である。

タスクが完了すると、タスクとスクリプトは削除され、対象のリモートホストに痕跡は残らない。

AutoReboot スイッチが有効な場合、Windows Update が完了すると、再起動要否が確認され、
必要であれば対象のリモートホストは再起動される。

最後に、利用可能な Update の有無が確認される。利用可能な Update が無い場合、リモート
ホストに適用されている Update の一覧がローカルホストのカレントディレクトリに出力される。
ジョブとして実行されている場合は、完了ステータスが、Completed になる。
利用可能な Update が残っている場合、ローカルホストのカレントディレクトリにその一覧が
出力され、アップデート未完了の例外が発生する。従って、ジョブとして実行されている場合は、
完了ステータスが Failed になる。

完了ステータスが Failed になった場合は、単純に Invoke-WindowsUpdate を再実行するだけ
で、たいてい解決する。問題がある場合はトラブルシューティングの項目を参照のこと。


# 使い方

Import-Module -Verbose -Force RemoteWindowsUpdate.psm1
Invoke-WindowsUpdate -AsJob -AutoReboot -ComputerName {対象ホスト名}


# 利用要件

Windows Server 2012R2 以上
もしくは PowerShell 3.0 以上で WinRM が利用可能であること


# トラブルシューティング

Invoke-WindowsUpdate を実行すると、すぐに異常終了してしまう場合、同名のタスクが
実行中である可能性がある。先行しているタスクが終了するのを待ってから再実行する。

前回のスクリプトが異常終了し、一時ファイルが残存した場合も同じ問題が発生する。
その場合は、リモートホストに作成された C:\Invoke-RemoteTask.ps1 を削除することで解消する。
