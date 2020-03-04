<#
    .SYNOPSIS
    archdoc
    Version: 0.03 27.02.2019
    
    © Anton Kosenko mail:Anton.Kosenko@gmail.com
    Licensed under the Apache License, Version 2.0

    .DESCRIPTION
    This script archive files with specific extension
#>

# requires -version 3

# Log function
    $Logfile = ".\info.log"
    Function LogWrite
    {
        Param ([string]$logstring)
        Add-content $Logfile -value $logstring
    }
# Declare Variable
    $Datestamp = Get-Date -Format "yyyyMMdd"
    $Directories=Import-Csv .\Directories.csv
# Set to zero counters
    $CntFlBef=0
    $CntFlAft=0
    $CntFoldBef=0
    $CntFoldAft=0
    $CntZip=0
    $Mail_TextBody = ""
# Mail sending function
    function MailError
    {
        $PSEmailServer = "ip or domain name"
        $Mail_HTMLBody = "<head><style>table {border-collapse: collapse; padding: 2px;}table, td, th {border: 1px solid #ffffff;}</style></head>"
        $Mail_HTMLBody += "<body style='background:#ffffff'><font face='Courier New'; size='2' color=#cc0000>"		
        $Mail_Subject = "ArchDoc $Datestamp"
        $Mail_HTMLBody += "<center><h2>ArchDoc Error Report</h2></center>"
        $Mail_HTMLBody += $Mail_TextBody
        $Mail_HTMLBody += "</font></body>"
        Send-MailMessage -From "service account mail" -To "admin mail" -Subject $Mail_Subject -Body $Mail_HTMLBody -BodyAsHtml -Encoding UTF8
        }
# Check folder function
    function Check {
        $ChDir=Test-Path $Folder
        if ($ChDir -match "False") {
            LogWrite "`t###### FAILED ######"
            LogWrite "`t$Folder not response or not exist" 
            $Mail_TextBody += "<center>$Folder are not response or not exist.</center>`n"
            MailError
            $StopScript=Get-Date
            $TimeToExec=($StopScript-$StartScript).Minutes
            LogWrite "`tScripts are executed $TimeToExec minutes"
            LogWrite "######## Abort $StopScript###################`n"
            exit
            }
            LogWrite "`tCheck Succes $Folder are exist and response"
        }
# Function moving files with specific extension
    function ArchivingFiles {
        $DirectoryTemp=Get-Childitem $Folder -Include "*.$Ext" -Recurse -Attributes !Directory | Where-Object LastWriteTime -le $Date
        $CntFlbef=$DirectoryTemp.Count
        if ($CntFlbef -gt 1) {
            $FolderName=$Folder+$Datestamp+$Ext
            if (!(test-path -path $FolderName)) {new-item -path $FolderName -itemtype directory | Out-Null}
            foreach ($files in $DirectoryTemp) {
                $FileName=$files.FullName
                Move-Item -Path $FileName -Destination $FolderName
            }
        }
        else {
            LogWrite "`tFiles are not exist"
            $Extension="False"
        }
    # Check moved files
        $CntFlAft=(Get-ChildItem $FolderName).Count
        if ($CntFlbef -eq $CntFlAft) {
            LogWrite "`tMove succes"
        # Archive temporary folder with target files and delete it
            $Archname=$Folder+$Datestamp+$Ext
                try {
                    Compress-Archive -Path $FolderName -CompressionLevel Optimal -DestinationPath  $Archname -Update -ErrorAction Stop
                    Remove-Item -Path $FolderName -Recurse
                    LogWrite "`tArchiving files completed"
                }
                catch {
                    LogWrite "`tMove Failed! Compress Error"
                    $Mail_TextBody += "<center>Archive Failed!</center>`n"
                    MailError
                    $StopScript=Get-Date
                    $TimeToExec=($StopScript-$StartScript).Minutes
                    LogWrite "`tScripts are executed $TimeToExec minutes"
                    LogWrite "######## Abort $StopScript###################`n"
                    exit    
                }
        }
        else {
            if ($Extension -eq "False") {continue}
            LogWrite "`tMove Failed!  $CntFlAft files are not moving to folder"
            $Mail_TextBody += "<center>Archive Failed!</center>`n"
            MailError
            $StopScript=Get-Date
            $TimeToExec=($StopScript-$StartScript).Minutes
            LogWrite "`tScripts are executed $TimeToExec minutes"
            LogWrite "######## Abort $StopScript###################`n"
            exit    
        }       
    }    
# Function selecting folder older target month and archiving it
    function ArchivingFolders {
        $DirectorySrc=Get-Childitem $Folder -Attributes Directory | Where-Object LastWriteTime -le $Date
        $CntFoldBef=($DirectorySrc).Count
    # Archive folder
        if ($CntFoldBef -gt 1) {
            foreach ($folders in $DirectorySrc) {
                $CntZip=$CntZip+1
                $Archname=$Folder+$Datestamp+$Ext
                $ExistFiles=(Get-ChildItem $folders.FullName).Count
            # Move file to archive
                if ($ExistFiles -lt 1) {continue}
                    try {
                        Compress-Archive -Path $Folders.FullName -CompressionLevel Optimal -DestinationPath  $Archname -Update -ErrorAction Stop
                        Remove-Item -Path $Folders.FullName -Recurse
                    }
                    catch {
                        LogWrite "`tMove Failed! Compress Error"
                        $Mail_TextBody += "<center>Archive Failed!</center>`n"
                        MailError
                        $StopScript=Get-Date
                        $TimeToExec=($StopScript-$StartScript).Minutes
                        LogWrite "`tScripts are executed $TimeToExec minutes"
                        LogWrite "######## Abort $StopScript###################`n"
                        exit    
                    }
                }
        }
        else {
            LogWrite "`tArchive Failed! Folders are not exist"
        }
    # Check archive
        $CntFoldAft=(Get-Childitem $Folder -Attributes Directory).Count
        if (($CntFoldBef -eq $CntZip) -or ($CntFoldBef -eq $CntFoldAft))
            {
            LogWrite "`tSucces"
            }
        else {
                LogWrite "`tArchive Failed! $CntFoldAft folders are not adding to archive"
                $Mail_TextBody += "<center>Archive Failed!</center>`n"
                MailError      
                }
            }
# Run script
    $StartScript=Get-Date
    LogWrite "######## Start $StartScript ###################"
# Select folders
    foreach ($Directory in $Directories) {
        if (($null -eq $Directory) -or ($Directory -eq "")) { continue }
    # Select month
        $Month=$Directory.Month
        $Date=(Get-Date).AddMonths($Month) 
        $Folder=$Directory.Directory 
        $Ext=$Directory.Extension
        Check
        ArchivingFiles
        ArchivingFolders
        }
    $StopScript=Get-Date
    $TimeToExec=($StopScript-$StartScript).Minutes
    LogWrite "`tScripts are executed $TimeToExec minutes"
    LogWrite "######## End $StopScript###################`n"
