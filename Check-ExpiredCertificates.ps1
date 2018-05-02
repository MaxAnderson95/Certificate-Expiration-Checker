[CmdletBinding()]
Param (

    [Parameter(Mandatory=$True)]
    [ValidateScript({
        if(!($_ | Test-Path) ){
            throw "File or folder does not exist" 
        }

        if(!($_ | Test-Path -PathType Leaf) ){
            throw "The Path argument must be a file. Folder paths are not allowed."
        }

        if($_ -notmatch "(\.csv)"){
            throw "The file specified must be a .CSV file."
        }

        return $true
    })]
    [System.IO.FileInfo]$PathToCSV,

    [String]$LogFileName = "$PSScriptRoot\Check-ExpiredCertificates-Log.log"

)

Import-Module "$PSScriptRoot\Get-WebCertificate.psm1"

Function Write-Log {

  Param (
    
    [String]$Message,
    
    [String]$LogFile
    
  )
  
  $Output = "$(Get-Date -Format G): $Message"
  
  Add-Content -Value $Output -Path $LogFile

}

Write-Log -Message "Starting Script" -LogFile $LogFileName

$CSV = Import-CSV -Path $PathToCSV -ErrorAction SilentlyContinue -ErrorVariable Error_ImportCSV

If ($Error_ImportCSV) {

    Write-Log -Message "Error importing CSV file." -LogFile $LogFileName

}

ForEach ($Site in $CSV) {

    Write-verbose "Contents of Site (1): $Site"
    
    Write-Verbose "Contents of Site (2): $Site"
    
    $Certificate = Get-WebCertificate -FQDN $Site.FQDN -Port $Site.Port

    Write-verbose "Contents of Site (3): $Site"

    If ($Certificate -eq $Null) {

        Write-Verbose "Error pulling certificate for $($Site.FQDN)"
        Write-Log -Message "Error pulling certificate for $($Site.FQDN)" -LogFile $LogFileName
        Continue

    }

    If ($Certificate.NotAfter.Date -eq (Get-Date).AddDays($Site.AlertDaysBefore).Date) {

        $EmailParams = @{

            To = $Site.EmailContact
            From = ""
            Subject = "SSL Certificate Expiring Soon"
            BodyAsHTML = $True
            Body = "The certificate for <b>$($Site.FQDN)</b> will be expiring in <b>$($Site.AlertDaysBefore) days</b>. Please renew this certificate soon."
            SMTPServer = ""
        
        }
        
        Send-MailMessage @EmailParams
        
        Write-Log -Message "$($Site.FQDN)'s certificate is expiring in $($Site.AlertDaysBefore) days. Email sent to $($Site.EmailContact)." -LogFile $LogFileName

    }

    Else {

        Write-Log -Message "$($Site.FQDN)'s certificate is not expiring in $($Site.AlertDaysBefore) days. No email was sent." -LogFile $LogFileName

    }

}

Write-Log -Message "Exiting Script" -LogFile $LogFileName