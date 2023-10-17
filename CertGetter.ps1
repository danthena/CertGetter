## Input files and initial var declarations

# Input file location
$ServersOutputlocation = "C:\temp\CertGetterTestedServers.txt"
$ADServersOutputLocation = "C:\temp\CertGetterADServers.txt"
$OfflineServersOutputLocation = "C:\temp\CertGetterADServersOffline.txt"

# Output for powershell transcript session
$TranscriptOutputLocation = "C:\temp\CertGetterTranscript.txt"

# Gets the short domain name for output files
$Domain = (net config workstation) -match 'Workstation domain\s+\S+$' -replace '.+?(\S+)$','$1'

# Counters
$OfflineServerCount = 0
$OnlineServerCount = 0
$CertsCollected = 0
$CurrentServerCount = 0
$TotalCertsCollected = 0

$ErrorActionPreference = 'Continue'

#Empty Arrays for later
$Array = @()
$Array2 = @()
$Array3 = @()
 
# Start session recording for errors
Start-Transcript -Path $TranscriptOutputLocation -force

# Bypass the execution policy warning message
Set-ExecutionPolicy Bypass -Force

#Create ADSservers.txt file
Write-Host Building initial server list... -ForegroundColor Yellow

#Create and populate CertGetterADSservers.txt file
Get-ADComputer -filter {OperatingSystem -like "*windows*server*"} | Select -Expand Name | Sort | sc $ADServersOutputLocation -Force

Write-Host Servers retreived! -ForegroundColor Green
Write-Host Starting connectivity tests... -ForegroundColor Yellow

$ADServers = Get-Content $ADServersOutputLocation

    foreach($ADServer in $ADServers)
    {
    
	#Checking to see if the computer is offline	
    If (!(Test-Connection -ComputerName $ADServer -Count 1 -Quiet)) { 
  
  Write-Host $ADServer is Unreachable -ForegroundColor Red
  $OfflineServerCount++
  $Array3 += $ADServer
  
  Continue # Move to next computer
	}
        else
        {
        $Array2 += $ADServer

        Write-Host $ADServer is Online -ForegroundColor Green
		$OnlineServerCount++
        }
		
    }
if($Array2)
{
    $Array2 | Out-File -FilePath $ServersOutputLocation -Force
}
    else
	{
    Write-Warning "ADServers checking failed"
    }
	
if($Array3)
{
    $Array3 | Out-File -FilePath $OfflineServersOutputLocation -Force
}
    else
	{
    Write-Warning "ADServers checking failed"
    }

 Write-Host Server Discovery Complete! $OfflineServerCount servers were Unreachable. $OnlineServerCount servers Online. Starting SSL scans on $OnlineServerCount servers -ForegroundColor Yellow
  
# Initilize the file and run the loops
$Servers = Get-Content $ServersOutputlocation
 
foreach($Server in $Servers)

{   
    $CertsCollected = 0
	
	Write-Host Scanning $Server -ForegroundColor green
	    
	# Adding this to a try/catch because something's bound to throw an error.
    Try
    {
        # Checking for hostname of server provided in input file 
        $hostname = ([System.Net.Dns]::GetHostByName("$Server")).hostname
   
        # Looking for those juicy certs
		$Certs = Invoke-Command $Server -ScriptBlock{ Get-ChildItem Cert:\LocalMachine\My }
		
		$CurrentServerCount++		
		
    }
    Catch
    {
        # Error message explaining what happened...
		$_.Exception.Message
        Continue
    }
      
    If($hostname -and $Certs)
    {
        Foreach($Cert in $Certs)
        {
            # Adding certificate properties and server names to an object
            $Object = New-Object pscustomobject 
            $Object | Add-Member Noteproperty "Server name" -Value $hostname
			# Out-String.Trim() was needed below to convert the system object to a string for the Csv 
            $Object | Add-Member Noteproperty "Certificate name" -Value ($cert.dnsnamelist.punycode | Out-String).Trim()
			      $Object | Add-Member Noteproperty "Certificate friendly name" -Value $cert.friendlyname
			      $Object | Add-Member Noteproperty "Certificate subject" -Value $cert.subject
            $Object | Add-Member Noteproperty "Certificate issuer"  -Value $cert.issuer
			      $Object | Add-Member Noteproperty "Certificate start date" -Value $cert.notbefore
            $Object | Add-Member Noteproperty "Certificate expiration date" -Value $cert.notafter  
            $Object | Add-Member Noteproperty "Certificate thumbprint" -Value $cert.thumbprint
			
			
			        if ($Cert.RawData){


				        # Public key to Base64
				        $CertBase64 = [System.Convert]::ToBase64String($Cert.RawData, [System.Base64FormattingOptions]::InsertLineBreaks)

				        # Put it all together
				        $Pem = @"
-----BEGIN CERTIFICATE-----
$CertBase64
-----END CERTIFICATE-----
"@

				  $Object | Add-Member Noteproperty "Certificate Encoded" -Value $Pem.ToString()
			}
			else {
				continue
				}  
   
            # Adding that object to an array
            $Array += $Object
			$CertsCollected++
			$TotalCertsCollected++
		}
    } 
    Else
    {
        Write-Warning "Uh oh..."
    }
	Write-Host $CurrentServerCount of $OnlineServerCount servers tested. $CertsCollected SSLs collected. -ForegroundColor Yellow
}
 
If($Array)
{
    # Export to CSV
    $Array | Export-Csv -Path C:\temp\CertGetterResults$Domain.csv -Force -NoTypeInformation
    Write-Host $TotalCertsCollected certificates retreived from $OnlineServerCount servers! -Foregroundcolor Green
}

Else
{
	Write-warning "Array Creation failed. If you're seeing this, something terrible happened or maybe you just need to check the permissions / location of the output file"
}

Stop-Transcript
