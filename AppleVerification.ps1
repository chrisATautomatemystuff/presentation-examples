<#
Author/Cobbler: Chris Thomas
Organization: Bloomfield Hills Schools
Email Address: chthomas@bloomfield.org

Creation Date: 10/01/14
Last Revision: 10/06/14

This script was cobbled together in order to verify Apple ID's in preparation for a large iPad rollout.
It draws from the scripts and ideas of others and uses three modules you'll likely have to download.

You'll want a CSV with a header row of:

    username,exchangepwd,appleidpwd

It will import the Exchange email address and password from your CSV (edit on lines 51-55), then
connect with Exchange Web Services (EWS) to the Client Access Server (CAS) URL (edit on line 105) and
pull the verfication URL from the Inbox, then use AutoBrowse to open an Internet Explorer browser session
with that verification URL, then tab around using WASP and login as the same Apple ID email address and
the Apple ID password from your CSV (edit on lines 51-55).

I'm sure there are more elegant ways to accomplish this, but ferme la bouche. ;-)
 
INSPIRATION: Communication with Apple iTunes Store and WebSite
http://d-fens.ch/2013/04/28/communication-with-apple-itunes-store-and-website/

MAIN "BORROWED" SCRIPT: Parsing out URL's in the body of a Message with EWS and Powershell
http://gsexdev.blogspot.com/2013/10/parsing-out-urls-in-body-of-message.html

EWS MODULE: Microsoft Exchange Web Services Managed API 2.0 (edit on line 48)
http://www.microsoft.com/en-us/download/details.aspx?id=35371

AUTOBROWSE MODULE: AutoBrowse - automate even the most annoying webpage (edit on line 45)
http://autobrowse.start-automating.com/
http://gallery.technet.microsoft.com/AutoBrowse-ec4f4384

WASP MODULE: Windows Automation Snapin for PowerShell (edit on line 42)
http://wasp.codeplex.com/
#>

## Import WASP
Import-Module 'C:\Users\chthomas\Documents\WindowsPowerShell\Modules\WASP\WASP.dll'

## Import AutoBrowse
Import-Module 'C:\Users\chthomas\Documents\WindowsPowerShell\Modules\AutoBrowse\AutoBrowse.psm1'

## Load EWS API
Add-Type -Path 'C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll'

## Define batch variables
$batch = "batch1"
$batchImport = "C:\appleids_" + $batch + ".csv"

$batchSuccess = "C:\appleids_" + $batch + "_success.csv"
$batchFailure = "C:\appleids_" + $batch + "_failure.csv"

## Get the Mailbox to Access from the 1st commandline argument
$MailboxName = $args[0]

## Set Exchange Version
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2

## Create Exchange Service Object
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

## Set Credentials to use
$psCred = @(Import-Csv $batchImport) | ForEach-Object {

$creds = New-Object System.Net.NetworkCredential($_.username.ToString(),$_.exchangepwd.ToString())
$service.Credentials = $creds

## Choose to ignore any SSL Warning issues caused by Self Signed Certificates
## Code From http://poshcode.org/624
## Create a compilation environment
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Compiler=$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() {
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert,
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

## We now create an instance of the TrustAll and attach it to the ServicePointManager
$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll
## end code from http://poshcode.org/624

## Set the URL of the CAS (Client Access Server) to use
$uri=[system.URI] "https://<FQDN of Exchange>/ews/exchange.asmx"
$service.Url = $uri

## Optional section for Exchange Impersonation
$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
# Bind to the Inbox Folder
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

#Define ItemView to retrive just 1 Item
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)
$fiItems = $service.FindItems($Inbox.Id,$ivItemView)
[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)
foreach($Item in $fiItems.Items){
	#Process Item
	"Processing : " + $Item.Subject
	$dupChk = @{}
	$RegExHtmlLinks = "<a href=\`"(.*?)\`">"
	$matchedItems = [regex]::matches($Item.Body, $RegExHtmlLinks,[system.Text.RegularExpressions.RegexOptions]::Singleline)
	foreach($Match in $matchedItems){
	    $SplitVal = $Match.Value.Split('"')
	    if($SplitVal.Count -gt 0){
	        $ParsedURI=[system.URI]$SplitVal[1]
			if($ParsedURI.Host -eq "id.apple.com"){
				if(!$dupChk.Contains($ParsedURI.AbsoluteUri)){
					#Write-Host -ForegroundColor Green	"AppleURL 	 : " + $ParsedURI.AbsoluteUri
					$dupChk.add($ParsedURI.AbsoluteUri,0)
                    
                    #Define the Apple ID email address and Apple ID password
                    $appleID = $_.username.ToString()
                    $appleIDPwd = $_.appleidpwd.ToString()

                    #Open a browser session with the Apple ID verification URL
                    Open-Browser -Url "$ParsedURI.AbsoluteUri" -Visible
                    $ie = Select-Window IEXPLORE | Select -First 1 | Set-WindowActive
                    Start-Sleep 5
                    
                    #Scrape site and skip the tabs if the Apple ID has already been verified
                    $p = Invoke-WebRequest "$ParsedURI"
                    
                    if($p.ParsedHtml.body.outerText -like "*has already been verified*"){
                        
                        $alreadyBeenVerified = "$appleID,alreadyBeenVerified"
                        $alreadyBeenVerified | Out-File $batchSuccess -Append
                        Write-Host $alreadyBeenVerified

                        #Close the browser before the next Apple ID is loaded
                        Select-Window IEXPLORE | Remove-Window
                    }
                    else{
                                            
                        $ie = Select-Window IEXPLORE | Select -First 1 | Set-WindowActive
                        Start-Sleep -Milliseconds 500

                        #Press Alt+D to ensure focus starts in the address bar
                        $ie | Send-Keys "%d"
                        Start-Sleep -Milliseconds 500

                        #Tab 26 times "because reasons"
                        1..26 | % {
                            $ie | Send-Keys "{TAB}"
                            Start-Sleep -Milliseconds 250
                            }

                        #Enter the Apple ID email address and password
                        $ie | Send-Keys "$appleID"
                        Start-Sleep -Milliseconds 500
                        $ie | Send-Keys "{TAB}"
                        Start-Sleep -Milliseconds 500
                        $ie | Send-Keys "$appleIDPwd"
                        Start-Sleep -Milliseconds 500
                        $ie | Send-Keys "{ENTER}"
                        Start-Sleep 5

                        #Check if the verification succeeded, display it on screen and log it
                        $verifiedURL = $ie | Get-Browser | Select -expand LocationURL
                    
                        if($verifiedURL -like "https://id.apple.com/IDMSEmailVetting/authenticate.html*"){
                            
                            $hasBeenVerified = "$appleID,hasBeenVerified"
                            $hasBeenVerified | Out-File $batchSuccess -Append
                            Write-Host $hasBeenVerified

                            #Close the browser before the next Apple ID is loaded
                            Select-Window IEXPLORE | Remove-Window
                            
                        }
                        elseif($verifiedURL -eq $ParsedURI){
                            $somethingWentWrong = "Something went wrong with $appleID"
                            $somethingWentWrong | Out-File $batchFailure -Append
                            Write-Host $somethingWentWrong

                            #Close the browser before the next Apple ID is loaded
                            Select-Window IEXPLORE | Remove-Window
                        }
                        #Check if the tabs messed up again, display it on screen and log it
                        else{
                            $somethingWentWrong = "$appleID,failure"
                            $somethingWentWrong | Out-File $batchFailure -Append
                            Write-Host $somethingWentWrong

                            #Close the browser before the next Apple ID is loaded
                            Select-Window IEXPLORE | Remove-Window
                        }
                    }
				}
			}
	    }
	}
}
}