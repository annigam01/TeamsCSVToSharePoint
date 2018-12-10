param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
    [string]$SharePointsiteURL = "", 
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
    [string]$ListTitle = "",
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
    [string]$username = "",
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
    [string]$password = "",
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
    [string]$CSVFile = ""
)

function Log([string] $msg)
{
$filpath = $PSScriptRoot+"\PSLog.log"

$msg = (Get-Date -Format g).ToString() +" " +$msg;
$msg | Out-File -FilePath $filpath -Append

}

function CreateSPList(){

log("creating list");

 #$exename = $PSScriptRoot+"\myexe.exe";
 $exename = "C:\Users\annigam.FAREAST\source\repos\callfromps1\myexe\CreateList\bin\Debug\CreateList.exe";

 $params9 = $SharePointsiteURL
 $params10 = $username
 $params11 = $password
 $params12 = $ListTitle
 
 &$exename $params9 $params10 $params11 $params12
 log("Finished creating list");
 }
function ReadCSV(){
 
 log("Reading CSV");

 try{
    $CSVItemCollection = Import-Csv -Path $CSVFile

     foreach ($CSVItem in $CSVItemCollection)
     {
          CallExe $CSVItem
          
     }
      log("Reading CSV");
 }
 catch{

    log($_.Exception.Message);
 }


 }
function CallExe($CsvObject){
log("Working on ")
log($CsvObject)

 #$exename = $PSScriptRoot+"\myexe.exe";
 $exename = "C:\Users\annigam.FAREAST\source\repos\callfromps1\myexe\myexe\bin\Debug\myexe.exe";

 $params1 = $CsvObject.From
 $params2 = $CsvObject.Team
 $params3 = $CsvObject.Channel
 $params4 = $CsvObject.Thread
 $params5 = $CsvObject.Subject
 $params6 = $CsvObject.Timestamp
 $params7 = $CsvObject.WebClientReadURL
 $params8 = $CsvObject.Message

 $params9 = $SharePointsiteURL
 $params10 = $username
 $params11 = $password
 $params12 = $ListTitle

 &$exename $params1 $params2 $params3 $params4 $params5 $params6 $params7 $params8 $params9 $params10 $params11 $params12
 log("Finished working")
}

log("script finished")
CreateSPList
ReadCSV
log("script finished")
