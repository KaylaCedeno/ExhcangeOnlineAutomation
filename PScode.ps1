Install-Module -Name ImportExcel

// incase settings don't allow you to run a script (& import libraries)

Set-ExecutionPolicy RemoteSigned 
Import-Module ImportExcel


// Importing the ExchangeOnline library

Install-Module -Name ExchangeOnlineManagement
Connect-ExchangeOnline


// Accessing excel sheet 

$ExcelFile = "C:\Users\KaylaCedeno\Downloads\IT onboarding list COPY.xlsx"
$Sheet = Import-Excel -Path $ExcelFile

// Iteration

  foreach ($col in $Sheet) {

    $access = $col.'Needs FF acess'
    $account = $col.'Account'


      // if "YES" and there is an account in that same row, then give them the fresh force test attribute

      if($access -eq 'YES' -and $account) {             
         Set-Mailbox $account -CustomAttribute2 "Fresh Force Test"
           }
       }
