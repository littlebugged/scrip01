######

按KB号安装补丁


#######


$sysInfo = Get-WmiObject -Class win32_OperatingSystem
If ($sysInfo.Version -contains '6.1.7601' -and $sysInfo.ProductType -contains '1'){
$ws = New-Object -ComObject WScript.Shell
$wsr=$ws.popup("是否进行补丁升级操作，点击 `“确定`” 立即开始,升级过程中请保证网络连通，本次升级持续40-60分钟，同时系统可能将会进行多次重启操作！！",0,"重要提示",1 + 64)
if($wsr -eq 2){
$wsr=$ws.popup("已取消升级，下次开机时将继续提示！！",0,"重要提示",1 + 48)
}
else{
#下载路径
If ($sysInfo.OSArchitecture -like "64*"){
		$source_Path = "http://192.168.0.100/kb/X64/"
	   }
	Else { $source_Path = "http://192.168.0.100/kb/X86/" }
$local_Path = "C:\DIR_KB\"
Write-Warning "补丁安装过程中请勿手动关闭窗口，同时系统可能将会进行多次重启操作，请提前保存好相应数据！"
sleep 30
#检查KB是否安装完成
function check_KB_install{
    #收集未安装成功的KB
    [System.Collections.ArrayList]$arraylist_kb=@()

    #获取当前已安装的KB
    $installedKBs_now=Get-HotFix

    #检查KB号是否安装完成
    foreach ($kb_now in $input){
        if(!($installedKBs_now -match $kb_now)){
            $arraylist_kb+=$kb_now        
        }
    }
    if($arraylist_kb.Count -eq 0){
        return "补丁已全部安装完成！"        
    }else{
        return 
    }        
}
#脚本路径
$script_Path = Split-Path -Parent $MyInvocation.MyCommand.Definition
#开机启动路径
$startup_folder = "$env:appdata" + "\Microsoft\Windows\Start Menu\Programs\Startup"
#补丁KB
$KBs = @("KB4490628", "KB4474419", "KB4536952", "KB4538483", "KB4537829", "KB4541110", "KB4537767", "KB4537813", "KB4541500")
$restart_kbs = @("KB4474419", "KB4538483", "KB4541500")
#调用函数检查KB号是否安装完成
[array]$fail_KB = $KBs|check_KB_install
#定义下载模块变量

$webClient = New-Object System.Net.webClient
#删除下载补文件
If ((Test-Path $local_Path)){
		Out-null -InputObject (Remove-Item -Path $local_Path -Recurse  -Force )
    }
#删除开机启动项
If (Test-Path "$startup_folder\install_KB.bat"){
    Out-Null -InputObject( Remove-Item -Path "$startup_folder\install_KB.bat" -Force)
    }
#已安装补丁列表
$installedKBs=Get-HotFix
#判断KB是否在补丁列表中
foreach($KB in $KBs){
if($installedKBs -match $KB){
Write-Host "$fail_KB" "$KB 已安装" -ForegroundColor Green
sleep 2
}
else{
$web_path = "$source_Path" + "$KB" + ".zip"
$local_KBzip = "$local_Path" + "$KB" + ".zip"
$local_KBmsu = "$local_Path" + "$KB" + ".msu"
#创建补丁存放路径
Out-null -InputObject (New-Item -Name DIR_KB -Path C:\ -ItemType directory -Force)
#下载并解压补丁
$webClient.DownloadFile($web_path, $local_KBzip)
$Shell = New-Object -ComObject shell.application
		$Zip = $Shell.NameSpace("$local_KBzip")
		foreach ($item in $zip.items()){
			$Shell.NameSpace("$local_Path").copyHere($item)
		}
#删除zip文件（感觉没有必要性）
		Remove-Item $local_KBzip -Force
#判断KB是否需要重启
        If ($restart_kbs -match $KB){
        
            Write-Warning "系统即将重启！" 
            #安装补丁
			wusa.exe $local_KBmsu /quiet /forcerestart 
            $WU = Get-Process | Where-Object { $_.ProcessName -like 'wusa' }
            #创建开机启动项
            Out-Null -InputObject( New-Item -Path $startup_folder -ItemType file -Name install_KB.bat -Force )
            Add-Content -Path "$startup_folder\install_KB.bat" -Value "Powershell -ExecutionPolicy Bypass $script_Path\install_KB_v2.ps1" -Force
            $WU.WaitForExit()
           #没有必要性
           Remove-Item $local_Path\*.msu
		    }
        Else
		  {
			wusa.exe $local_KBmsu /quiet /norestart
			$WU = Get-Process | Where-Object { $_.ProcessName -like 'wusa' }
			$WU.WaitForExit()
			Remove-Item $local_Path\*.msu
		  }
}
}
    }
    }
Else { "Non-windows 7 SP1 system, script is about to exit!!" }
