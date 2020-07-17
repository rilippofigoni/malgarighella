##most of this code was from Glen Scales on http://social.msdn.microsoft.com/Forums/exchange/en-US/e8af16ac-845d-4afb-a14a-3c373a80e932/how-to-export-calendar-items-##to-csvfile-with-powershell?forum=exchangesvrdevelopment&prof=required
##I also used some from http://msgdev.mvps.org/exdevblog/expapt.zip - also by Glen Scales. The zipped file has a PS script.



## Get the Mailbox to Access from the 1st commandline argument, put in the target email address here (make sure you have permissions)

$MailboxName = "user@domain.ie"

## Load Managed API dll  
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"  
  
## Set Exchange Version  
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2  
  
## Create Exchange Service Object  
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
  

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString()) 
  
  
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
  
## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
  
#CAS URL Option 1 Autodiscover  
$service.AutodiscoverUrl($MailboxName,{$true})  
"Using CAS Server : " + $Service.url   
   
#CAS URL Option 2 Hardcoded  
  
#$uri=[system.URI] "https://exchangeServer/ews/exchange.asmx"  
#$service.Url = $uri    
  
## Optional section for Exchange Impersonation  
  
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
#Bind to Calendar    
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)     
$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)    
  
$exportCollection = @() 

#Define Date to Query - this involves manually entering the start & end date
$StartDate = $(Read-Host "Enter Start day of Calendar in format YYYY-MM-DD")
$EndDate = $(Read-Host "Enter End day of Calendar in format YYYY-MM-DD") 
  
#Define the calendar view  
$CalendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($StartDate,$EndDate,1000)    
$fiItems = $service.FindAppointments($Calendar.Id,$CalendarView)    
foreach($Item in $fiItems.Items){      
	$exportObj = "" | Select StartTime,EndTime,Subject,Location
    "Start    : " + $Item.Start  
    "Subject  : " + $Item.Subject  
	$exportObj.StartTime = $Item.Start
	$exportObj.EndTime = $Item.End
	$exportObj.Subject = $Item.Subject
	$exportObj.Location = $Item.Location
	$exportCollection +=$exportObj
}
$FileName = (Get-Location).Path.ToString() + "\Calendar-" + (Get-Date).ToString("yyyy-MM-dd-hh-mm-ss") + ".csv"
$exportCollection | Export-Csv -NoTypeInformation -Path $FileName
