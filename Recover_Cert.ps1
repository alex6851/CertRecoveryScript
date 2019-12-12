

function Format-Path{
    Param ([string]$ExportPath)
    $ExportPath = $ExportPath.Trim()
    If (!( $Exportpath -match '.+\\$')) 
    {
      $Exportpath += "\"
    }
    return $ExportPath
}
  
function Test-Copy {
  Param ( [string]$copypath)
  $testfile = "zztest-filexyz.txt"
  $testfilepath = "C:\Program Files\Cert Recover Script\zztest-filexyz.txt"
  $copydest = $copypath + $testfile

  Copy-Item $testfilepath $copydest -ErrorAction SilentlyContinue
  if ($?)
  {return $true}
  else 
  {return $false}
}

function Find-User {
  param (
      [string]$UserName
  )
  do {
      $UserExists = Get-ADUser -LDAPFilter "(sAMAccountName=$UserName)"
      If ($Null -eq $UserExists) {$UserName = Read-Host "Username $UserName is not valid or does not exist. Please specify a different username."}
      else {return $UserName}
  }
  Until ($Null -ne $UserExists)   
}

function Get-UDrivePath {
  param (
       $username
  )
$Udrive = Get-Aduser $username -properties HomeDirectory | Select HomeDirectory
return $Udrive.HomeDirectory
}

function New-CertFolder {
  param (
      $Exportpath
  )   

$certfolder = "not finished"
$Exportpath = Format-Path $Exportpath 
$Exportpathcertsfolder = $Exportpath + "CERTS\"
$Exportpathcertfolder = $Exportpath + "CERT\"

if (!(Test-Path $Exportpathcertfolder) -and !(Test-Path $Exportpathcertsfolder)) {
  while ( $certfolder -ne "finished") {
      $UserInput = Read-Host "CERT folder doesnt exist in target location do you want to create one?"
      $UserInput = $UserInput.Trim()
      if (($UserInput -eq "Y") -or ($UserInput -eq "Yes")) {
          New-Item -ItemType Directory -Force -Path $Exportpathcertfolder
          $certfolder = "finished"
      }
      elseif (($UserInput -eq "N") -or ($UserInput -eq "No")) {
          $certfolder = "finished"
      }
      elseif ($certfolder -ne "finished") {
          continue
      }
  }
}

elseif ((Test-Path $Exportpathcertfolder) -and !(Test-Path $Exportpathcertsfolder)) {
      while ( $certfolder -ne "finished") {
          $UserInput = Read-Host "There is a CERT folder do you want to put it there?"
          $UserInput = $UserInput.Trim()
          if (($UserInput -eq "Y") -or ($UserInput -eq "Yes")) {
              Get-Item $Exportpathcertfolder
              $certfolder = "finished"
          }
          elseif (($UserInput -eq "N") -or ($UserInput -eq "No")) {
              $certfolder = "finished"
          }
          elseif ($certfolder -ne "finished") {
              continue
          }
      }
  }

  elseif (!(Test-Path $Exportpathcertfolder) -and (Test-Path $Exportpathcertsfolder)) {
      while ( $certfolder -ne "finished") {
          $UserInput = Read-Host "There is a CERTS folder do you want to put it there?"
          $UserInput = $UserInput.Trim()
          if (($UserInput -eq "Y") -or ($UserInput -eq "Yes")) {
              Get-Item $Exportpathcertsfolder
              $certfolder = "finished"
          }
          elseif (($UserInput -eq "N") -or ($UserInput -eq "No")) {
              $certfolder = "finished"
          }
          elseif ($certfolder -ne "finished") {
              continue
          }
      }
  }
  elseif ((Test-Path $Exportpathcertfolder) -and (Test-Path $Exportpathcertsfolder)) {
      while ( $certfolder -ne "finished") {
          $UserInput = Read-Host "There is a CERT folder AND a CERTS folder do you want to put it in the CERT folder?"
          $UserInput = $UserInput.Trim()
          if (($UserInput -eq "Y") -or ($UserInput -eq "Yes")) {
              Get-Item $Exportpathcertfolder
              $certfolder = "finished"
          }
          elseif (($UserInput -eq "N") -or ($UserInput -eq "No")) {
              $certfolder = "finished"
          }
          elseif ($certfolder -ne "finished") {
              continue
          }
      }
  }

  
}



$UserName = Read-Host "Which user's cert do you want?"

$UserName = $UserName.Trim()

$UserName = Find-User $UserName

$output = certutil -view -restrict "-CommonName == $username" -out "RequesterName,CommonName,Certificate Expiration Date,SerialNumber" | Tee-Object -FilePath "C:\Program Files\Cert Recover Script\temp_regex2.txt"

$output

$SerialNumbers = @()
$dates = @()
ForEach ($object in $output)
{
  if ($object -match '\w{38}')
  {
      $object = $object.substring(17)
      $object = $object.replace('"',"")
      $SerialNumbers += $object
  }
  if ($object -match '\d+(\/)\d+(\/)\d{4}\s\d+:\d{2}\s\w{2}')
  {
      $object = $object.substring(31)
      $dates += $object
  }
}

$Row = Read-Host "Which cert do you want to recover? (Choose 1 -"$SerialNumbers.count")"
$Row = $Row - 1

$SerialNumber = $SerialNumbers[$Row]
$ExpirationDate = $dates[$Row]

$SerialNumber
$ExpirationDate


$ExpirationDate = $ExpirationDate -replace "(\/)", '.'
$ExpirationDate = $ExpirationDate -replace ":", '.'
$ExpirationDate = $ExpirationDate -replace "\s", '_'

$ExportFilename = $username + "_" + $ExpirationDate + ".pfx"

$SendtoUDrive = "not finished"
While ($SendtoUDrive -ne "finished") {
$UserInput = Read-Host "Do you want to Export to the users U: drive?"
if (($UserInput -eq "Y") -or ($UserInput -eq "Yes")) {
   $Exportpath = Get-UDrivePath $UserName
   $Exportpath = New-CertFolder $Exportpath
   $Exportpath = Format-Path $Exportpath
   $SendtoUDrive = "finished"
}

elseif (($UserInput -eq "N") -or ($UserInput -eq "No")) {
  $SendtoUDrive = "finished"
  $CreateOwnPath = $true
   }
   else 
   {continue}
}
if ($CreateOwnPath -eq $true) 
{
  $Exportpath = Read-Host "Where do you want to Export the cert?"

  $Exportpath = Format-Path $Exportpath

  while (!(Test-path $Exportpath)) {
      $Exportpath = Read-Host "Folder does not exist. Please enter a valid Path."

      while ((Test-Path $Exportpath) -and !(Test-Copy $Exportpath)) {      
          $Exportpath = Read-Host "You do not have write permissions to this folder please select a different folder"
      }

      if (!(Test-path $Exportpath)) {
          continue
      }

  }   
  $Exportpath = Format-Path $Exportpath
  $Exportpath = New-CertFolder $Exportpath
  $Exportpath = Format-Path $Exportpath
  
}


$ExportPath += $ExportFilename

if (Test-Path $Exportpath)
{
  ""
  ""
  Write-Host $Exportpath already Exists
  ""
  ""
  break
}




$path = "C:\ITStuff\certs\Exports\$username\"
$folderWithDate = $path + $ExpirationDate



$input_path_regex = ‘C:\Program Files\Cert Recover Script\temp_regex.txt’
$output_file_batch = ‘C:\Program Files\Cert Recover Script\temp_script.bat’

echo "mode con: cols=299 lines=30" | out-file $output_file_batch -encoding ASCII

If (!(test-path $path)) {
  New-Item -ItemType Directory -Force -Path $path
  Set-Location $path
  echo "cd $path" | out-file $output_file_batch -Append -encoding ASCII
}
ElseIf (test-path $folderWithDate) {
  Remove-Item –path $folderWithDate –recurse -Force
  New-Item -ItemType Directory -Force -Path $folderWithDate
  Set-Location $folderWithDate
  echo "cd $folderWithDate" | out-file $output_file_batch -Append -encoding ASCII
}
Else {
  New-Item -ItemType Directory -Force -Path $folderWithDate
  Set-Location $folderWithDate
  echo "cd $folderWithDate" | out-file $output_file_batch -Append -encoding ASCII
}


certutil -getkey $SerialNumber > $input_path_regex


$RegexScriptBuilder = ‘certutil -.*’
#$KeyRegex = '"................."$'
$KeyRegex = '(?<=@echo PASSWORD: )".*"'
select-string -Path $input_path_regex -Pattern $RegexScriptBuilder -AllMatches | % { $_.Matches } | % { $_.Value } | out-file $output_file_batch -Append -encoding ASCII
#echo "cmd.exe" | out-file $output_file_batch -Append -encoding ASCII

start $output_file_batch -Wait

$key = select-string -Path $input_path_regex -Pattern $KeyRegex -AllMatches | % { $_.Matches } | % { $_.Value }

$ImportFilename = $username + ".p12"

certutil -f -p $key -importpfx $ImportFilename

certutil -exportpfx -p "Password you Create" $SerialNumber $ExportPath
  
  $UserInput = Read-Host "Do you want to E-mail the cert?"    
  switch -Wildcard ( $UserInput )
  {
      Y
      {
         $To = Read-Host "Who do you want to mail it to?"
         $Body = Read-Host "What do you want to say?"
         $Confirm = Read-Host "Do you want to send it?"          
          switch -Wildcard ( $Confirm )
         {
             Y
             {
                  $From = "alexander.baker@mrcy.com"
                  $SMTPServer = "Mail Server IPAddress"
                  $Subject = "$username's Certificate"
                  Send-mailmessage -To $To -From $From -SmtpServer $SMTPServer -Subject $Subject -Body $Body -Attachments $ExportPath
                  break
             }
             N
             {
                 break
             }
          }
      }
      N
      {
         break
      }
  }

break