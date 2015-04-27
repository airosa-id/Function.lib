### Общая библиотека функций
### Описание обновления версий: первая цифра - изменение/добавление функций, вторая - изменение/добавление переменных, третья - прочее
### Ver 5.4.0 - 2015-0330

Function CheckFile ($Filename) {
	$ShutdownTime = [datetime] "23:50:00"
	$encoding = [System.Text.Encoding]::UTF8
	while (!(Test-Path $FileName -PathType Leaf)){
		if ((Get-Date -DisplayHint  Time) -ge ($ShutdownTime)) {
		return $Exit_code = 1
		Exit
		}
		Start-Sleep 15
	}
	if (!(Test-Path $FileName -PathType Leaf)){
	}
	$ObjFile = Get-ChildItem $FileName
	if ($ObjFile.lastwritetime.Date -ne ([datetime]::Today)){
		$OwnerFile = (Get-Acl $ObjFile).Owner
		return $Exit_code = $ObjFile
		Exit
	} else {
		return $exit_code = 0
	}
}
### Функция запуска программ с аргументами и отслеживанием ошибки выхода
Function RunProc () {
	if (Test-Path $args[0]){
		$Index_arr = 0
		# Подсчет колличества аргументов функции за исключением первого. Потому что первый нифига не аргумент.
		foreach ($elements in $args) {
			$Index_arr += 1
			if ($Index_arr -eq 1) {continue}
			$arg_complete += " " + $elements
		}
		$pi = New-Object System.Diagnostics.ProcessStartInfo -Property @{
			"FileName" = $args[0]
			"Arguments" = $arg_complete
			"UseShellExecute" = $true
			"WorkingDirectory" = (Get-Location).providerpath
			"WindowStyle" = [System.Diagnostics.ProcessWindowStyle]::Hidden
		}
		$pr = [system.Diagnostics.Process]::Start($pi);
		$pr.WaitForExit();
		return $pr.ExitCode
	} else {
		$pr = "File not Exist"
		return $pr
	}
}
### Функция вычисления хэша MD5
function get_hash ($Filename) {
	if (Test-Path $Filename -PathType Leaf) {
		$hasher = [System.Security.Cryptography.MD5]::Create()
		$File_to_Stream = New-Object System.IO.StreamReader($Filename)
		$hash_result = $hasher.ComputeHash($File_to_Stream.BaseStream)
		$File_to_Stream.Close()
		$builder = New-Object System.Text.StringBuilder
	    $hash_result | Foreach-Object {[void]$builder.Append($_.ToString("X2"))} 
	    $output = New-Object PsObject 
	    $output | Add-Member NoteProperty HashValue ([string]$builder.ToString())
		return $output.hashvalue
	} else {
		$output = "Error. File not exist."
		return $output
	}
}

# Функция Записи в лог
function Log-out ($FLogPath=$LogPath,$Logname,$Text) {
		Out-File -Append -FilePath ($FLogPath+$(Get-Date -Format yyyy-MMdd)+"_"+$Logname) -InputObject ("$(Get-Date -DisplayHint DateTime)  $text")
}

#Функция Архивирования c помощью 7z
function Archive ([string]$Type="7z",$WorkDir=(Get-Location).providerpath,$ArchName,$FileNames) {
	Foreach ($FileName in $FileNames) {
		$ResultName = $ResultName + " " + $FileName
	}
	$Exit_code = (RunProc c:\"Program Files"\7-Zip\7z.exe a "-t$Type" "-w$WorkDir" $ArchName $ResultName)
	### Завершение работы архиватора
	switch ($Exit_code) {
			0	{$report_7z = "---------- Архивирование файлов..... ОК."}
			1	{$report_7z = "----------Warning (Non fatal error(s)). For example, one or more files were locked by some other application, so they were not compressed."}
			2	{$report_7z = "----------Fatal Error."}
			7	{$report_7z = "----------Command line error."}
			8	{$report_7z = "----------Not enough memory for operation."}
			255	{$report_7z = "----------User stopped the process."}
			default	{$report_7z = "----------Неизвестная ошибка: ("+$Exit_code+") – исполнение программы прервано."}
	}
	return $report_7z
}

#Функция распаковки файлов из архива в указанной директории
function Unpack ($Folder_with_Archive) {
	# Переходим в рабочий каталог аттачев
	Set-Location -Path $Folder_with_Archive
	# Перечитываем список файлов вложений, чистим от неправедных файлов, 
	# распаковываем найденные zip-файлы с их послед. разархивированием и удалением
	Log-out -Logname $LogName -Text "`n`n********** Распаковка и удаление zip-файлов. **********`n"
	$aFolder = Get-childitem $Folder_with_Archive
	if ($aFolder -ne $null) {
		ForEach ($aFile in $aFolder) {
			if ($aFile.Extension.ToLower() -ne ".zip" -and $aFile.Extension.ToLower() -ne ".p7e" -and $aFile.Extension.ToLower() -ne ".p7s" -and $aFile.Extension.ToLower() -ne ".p7a"-and $aFile.Extension.ToLower() -ne ".cry") {
				Log-out -Logname $LogName -Text ("Удаляем файл: " + $aFile.Name)
				Remove-Item $aFile.Name -Force
			}elseif ($aFile.Extension.ToLower() -eq ".zip") {
				Log-out -Logname $LogName -Text ("Распаковка: " + $aFile.Name)
				$Exit_code = (RunProc $Env:ProgramFiles\7-Zip\7z.exe x $($aFile.FullName) -aoa $("-o"+$Folder_with_Archive))
				switch ($Exit_code) {
					0		{Log-out -Logname $LogName -Text " 0 – OK"
					Log-out -Logname $LogName -Text ("Удаляем zip Файл: " + $aFile.Name)
							 Remove-Item $aFile.FullName -Force
							}
					default	{Log-out -Logname $LogName -Text " Ошибка: ("+$Exit_Code+")"}
				}
			}
		}
	} else {
		Log-out -Logname $LogName -Text "Пусто"
	}
}

# Перечитываем список файлов вложений и декриптуем *.p7e-файлы с их послед. удалением
function Decode_p7e ($Folder_with_p7e) {
	$aFolder = Get-childitem $Folder_with_p7e
	Set-Location -Path $Folder_with_p7e
	Log-out -Logname $LogName -Text "`n`n********** Расшифровываем и распаковываем p7e-файлы **********`n"
	if (!($aFolder -eq $null)) {
		ForEach ($aFile in $aFolder) {
			if ($aFile.Extension.ToLower() -eq ".p7e") {
				Log-out -Logname $LogName -Text ("Расшифровываем файл: " + $aFile.Name)
				$Exit_code = (RunProc $binFolder\xpki1utl.exe -decrypt -in $aFile.FullName -out $($aFile.DirectoryName+"\"+$($aFile.BaseName)) -silent ($LogPath+$(Get-Date -Format yyyy-MMdd)+"_Crypto_p7e.log") )
				switch ($Exit_code) {
					0{
						Log-out -Logname $LogName -Text " 0 – OK"
						# Тут же распаковываем архив
						Unpack ($Folder_with_p7e)
						# Удаляем p7e
						Remove-Item $aFile.FullName -Force
					}
					default	{
						Log-out -Logname $LogName -Text (" Ошибка расшифровки: ("+$Exit_Code+")")
					}
				}
			}
			if ($aFile.Extension.ToLower() -eq ".cry") {
				Log-out -Logname $LogName -Text ("Расшифровываем файл: " + $aFile.Name)
				$Exit_code = (RunProc $binFolder\xpki1utl.exe -decrypt -in $aFile.FullName -out $($aFile.DirectoryName+"\"+$($aFile.BaseName+".zip")) -silent ($LogPath+$(Get-Date -Format yyyy-MMdd)+"_Crypto_cry.log") )
				switch ($Exit_code) {
					0{
						Log-out -Logname $LogName -Text " 0 – OK"
						# Тут же распаковываем архив
						Unpack ($Folder_with_p7e)
						# Удаляем cry
						Remove-Item $aFile.FullName -Force
					}
					default	{
						Log-out -Logname $LogName -Text (" Ошибка расшифровки: ("+$Exit_Code+")")
					}
				}
			}
			if ($aFile.Extension.ToLower() -eq ".p7a") {
				Log-out -Logname $LogName -Text ("Расшифровываем файл: " + $aFile.Name)
				$Exit_code = (RunProc $binFolder\xpki1utl.exe -decrypt -in $aFile.FullName -out $($aFile.DirectoryName+"\"+$($aFile.BaseName)+".p7s") -silent ($LogPath+$(Get-Date -Format yyyy-MMdd)+"_Crypto_p7e.log") )
				switch ($Exit_code) {
					0{
						Log-out -Logname $LogName -Text " 0 – OK"
						# Удаляем p7a
						Remove-Item $aFile.FullName -Force
					}
					default	{
						Log-out -Logname $LogName -Text (" Ошибка расшифровки: ("+$Exit_Code+")")
					}
				}
			}
		}
	} else {
		Log-out  -Logname $LogName -Text "Зашифрованные файлы файлы отсутствуют"
	}
}

# Удаляем сигнатуры в *.p7s-файлах
function De_Signature ($Folder_with_p7s) {
	$aFolder = Get-ChildItem $Folder_with_p7s
	Set-Location -Path $Folder_with_p7s
	Log-out  -Logname $LogName -Text "`n`n********** Удаляем сигнатуры в *.p7s-файлах **********`n"
	if (!($aFolder -eq $null)) {
		ForEach ($aFile in $aFolder) {
			if ($aFile.Extension.ToLower() -eq ".p7s") {
				Log-out  -Logname $LogName -Text ("Удаление сигнатур: " + $aFile.FullName)
				$Exit_code = (RunProc $binFolder\xpki1utl.exe -verify -in $aFile.FullName -delete -1  -out $($aFile.DirectoryName+"\"+$($aFile.BaseName)) -silent ($LogPath+$(Get-Date -Format yyyy-MMdd)+"_Crypto_p7s.log"))
				switch ($Exit_code) {
					0		{Log-out -Logname $LogName -Text " 0 – OK"
							Log-out -Logname $LogName -Text ("Удаляем p7s Файл: " + $aFile.Name)
							 Remove-Item $aFile.FullName -Force
							}
					default	{Log-out -Logname $LogName -Text (" Ошибка: ("+$Exit_Code+")")}
				}
			}
		}
	} else {
		Log-out -Logname $LogName -Text "Пусто"
	}
}

Function AuthMailSender ($From,$To,$Subject,$Body,$File=$null,$UserName,$Password,$SMTPServer="mail.vivait-ic.ru",$SMTPPort="25") {
	$smtpClient = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort);
	$MailMessage = New-Object System.Net.Mail.MailMessage
	$smtpClient.Credentials = New-Object System.Net.NetworkCredential($Username, $Password);
	$smtpClient.EnableSSL = $true
	$MailMessage.IsBodyHTML = $true
	$MailMessage.Body = $Body
	$MailMessage.Sender = $From
	$MailMessage.From = $From
	if ($File) {
		$Attachment = new-object Net.Mail.Attachment($File)
		$MailMessage.Attachments.Add($Attachment)
	}
	$MailMessage.Subject = $Subject
	$MailMessage.To.add($To)
	$SMTPClient.Send($MailMessage)
}

### БЛОК ОПРЕДЕЛЕНИЯ ПЕРЕМЕННЫХ

$NDay = Get-Date -format dd
$NMonth = Get-Date -format MM
$Date_full = Get-Date -UFormat "%d.%m.%Y"
$Nyear = Get-Date -format yyyy
$NTime = Get-Date -Format T

# Подгрузка необходимых переменных для определенных скриптов
if (test-path variable:scriptName) {
	switch ($scriptName) {
		Mail_report.ps1 {
			#  каталог для извлеченных файлов
			$attachFolder = "D:\MailFolder\Attach\"
			#  каталог с утилитами
			$binFolder = "D:\_Script\Bin"
			# Хвост хранения архивов
			$history_delete = [datetime]::Today.AddDays(-14)
			# Каталог с логами
			$LogPath = "D:\_Script\_Logs\"
			# Имя лог файла по умолчанию
			$LogName = "mail.log"
			#  каталог с входящими письмами.
			$mailFolder = "D:\MailFolder\New\"
			#  каталог с отработанными письмами.
			$oldmailFolder = "D:\MailFolder\Old\"
			#  каталог для временных файлов.
			$tmpFolder = "D:\MailFolder\tmp\"
		}
		!TQServer_start.ps1 {
			# Путь до архива отстатоков и лимитов
			$ArchivePath = "\\vivait.lan\Company\_Public\Итоги_торгов\Saldos\Archive\"
			# Путь до шлюзов транзака
			$GatesPath = "D:\Transaq\Gates\"
			# Хвост хранения архивов
			$history_delete = [datetime]::Today.AddDays(-90)
			# Имя лог файла по умолчанию
			$LogName = "!TQServer_start.log"
			# Путь до логов самого скрипта загрузки
			$LogPath = "D:\Transaq\logs\StartScript\"
			# Дата для имени файла остатков вчерашнего дня
			$LTDday = (get-date).AddDays(-1).ToString("MMdd")
			# Путь до остатков на ММВБ и лимитов на FORTS
			$saldospath = "\\vivait.lan\Company\_Public\Итоги_торгов\Saldos\"
			# Путь до остатков на ММВБ Сгенерированных сервером транзака на конец дня
			$saldosLTDpath = "D:\Transaq\Report\"
			# Корневой путь до сервера транзака
			$RootPath = "D:\Transaq\Server\"
			# Путь до saldos.exe и pf.limits
			$UtilPath = "D:\Transaq\Utils\"
		}
	}
}
