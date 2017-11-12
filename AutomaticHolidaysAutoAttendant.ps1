### Automatic Holidays AutoAttendant            ###
### Version 1.0                                 ###
### Author: Alexander Holmeset                  ###
### Email: alexander.holmeset@gmail.com         ###
### Twitter: twitter.com/alxholmeset            ###
### Blog: alexholmeset.blog                     ###


#Enter what country you need holidays from. Se valid countries in Mostcountries function.
$Country = "Norway"


#Function for catching data from officeholidays.com, and convert it to a variable.
function Mostcountries {

        param(
        [Parameter(Position=0)]
        [ValidateSet("Algeria","Angola","Armenia","Argentina","Australia","Austria","Azerbaijan","Bahamas","Bahrain","Bangladesh","Barbados","Belarus","Belgium","Bolivia","Bosnia_and_Herzegovina","Botswana","Brazil","Brunei","Bulgaria","Burundi","Cambodia","Canada","Cayman_Islands","Chile","China","Colombia","Costa_Rica",
        "Croatia","Cyprus","Czech_Republic","Denmark","Dominican_Republic","Ecuador","Egypt","El_Salvador","Estonia","Ethiopia","Fiji","Finland","France","Georgia","Germany","Ghana","Gibraltar","Greece","Grenada","Guernsey",
        "Honduras","Hong_Kong","Hungary","Iceland","India","Indonesia","Iraq","Ireland","Isle_of_Man","Israel","Italy","Jamaica","Japan","Jersey","Jordan","Kazakhstan","Kenya","Kuwait","Lao",
        "Latvia","Lebanon","Libya","Liechtenstein","Lithuania","Luxembourg","Macau","Macedonia","Maldives","Malta","Mauritius","Mexico","Moldova","Monaco","Montenegro",
        "Malaysia","Morocco","Mozambique","Myanmar","Netherlands","New_Zealand","Nigeria","Norway","Oman","Pakistan","Panama","Paraguay","Peru","Philippines","Poland","Portugal","Qatar",
        "Romania","Russia","Rwanda","Saint_Lucia","Saudi_Arabia","Serbia","Singapore","Slovakia","Slovenia","South_Africa","South_Korea","Spain","Sri_Lanka","Sweden",
        "Switzerland","Taiwan","Tanzania","Thailand","Tonga","Trinidad_and_Tobago","Tunisia","Turkey","Turks_and_Caicos_Islands","Uganda","Uganda","Ukraine",
        "United_Arab_Emirates","United_Kingdom","Uruguay","USA","Venezuela","Vietnam","Yemen","Zambia","Zimbabwe")]
        [System.String]$Country
        )

$uri = "http://www.officeholidays.com/countries/$Country/index.php"
$html = Invoke-WebRequest -Uri $uri
$tables = $html.ParsedHtml.getElementsByTagName('tr') |
Where-Object {$_.classname -eq 'holiday' -or $_.classname -eq 'regional' -or $_.classname -eq 'publicholiday' } |
Select-Object -exp innerHTML
$script:holidays = foreach ($table In $tables){ 

$Date = (($table -split "<SPAN class=ad_head_728>")[1] -split "</SPAN>")[0]; 
$dateofmonth = $Date -replace '\D+(\d+)','$1'
$month = $Date -replace '[^a-zA-Z-]',''
$dayofmonth = $dateofmonth+$month

$Title = ((($table -split "<TD><A title=")[1] -split ">")[1] -split "</A")[0]
[PSCustomObject]@{
        Title = $Title ; Date = $dayofmonth | Get-Date -UFormat %d/%m/%Y
        }

 }
}





#


$ErrorActionPreference = "SilentlyContinue"




Mostcountries -country $Country


#Comment out from here, after first run.

#Set a availivle number here.
$lineUri = [System.Uri] "tel:+4799999999"

#who is going to be the operator for this AutoAttendant
$operatorUri = "sip:john@contoso.com"
$operatorEntity = New-CsOrganizationalAutoAttendantCallableEntity -Identity $operatorUri -Type User

$dcfGreetingPrompt = New-CsOrganizationalAutoAttendantPrompt -TextToSpeechPrompt "Welcome to Contoso!"
$dcfMenuOptionZero = New-CsOrganizationalAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone0
$dcfMenuPrompt = New-CsOrganizationalAutoAttendantPrompt -TextToSpeechPrompt "To reach your party by name, enter it now, followed by the pound sign or press 0 to reach the operator."
$dcfMenu=New-CsOrganizationalAutoAttendantMenu -Name "Default menu" -Prompts @($dcfMenuPrompt) -MenuOptions @($dcfMenuOptionZero) -EnableDialByName
$defaultCallFlow = New-CsOrganizationalAutoAttendantCallFlow -Name "Default call flow" -Greetings @($dcfGreetingPrompt) -Menu $dcfMenu


$afterHoursGreetingPrompt = New-CsOrganizationalAutoAttendantPrompt -TextToSpeechPrompt "Welcome to Contoso! Unfortunately, you have reached us outside of our business hours. We value your call please call us back Monday to Friday, between 9 A.M. to 12 P.M. and 1 P.M. to 5 P.M. Goodbye!"
$afterHoursMenuOption = New-CsOrganizationalAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic 
$afterHoursMenu=New-CsOrganizationalAutoAttendantMenu -Name "After Hours menu" -MenuOptions @($afterHoursMenuOption)
$afterHoursCallFlow = New-CsOrganizationalAutoAttendantCallFlow -Name "After Hours call flow" -Greetings @($afterHoursGreetingPrompt) -Menu $afterHoursMenu

#Defines and creates weekly operating hours for when its not a holiday.
$tr1 = New-CsOnlineTimeRange -Start 08:00 -End 16:00
$afterHoursSchedule = New-CsOnlineSchedule -Name "Business Hours" -WeeklyRecurrentSchedule -MondayHours @($tr1) -TuesdayHours @($tr1) -WednesdayHours @($tr1) -ThursdayHours @($tr1) -FridayHours @($tr1) -complement
$afterHoursCallHandlingAssociation = New-CsOrganizationalAutoAttendantCallHandlingAssociation -Type AfterHours -ScheduleId $afterHoursSchedule.Id -CallFlowId $afterHoursCallFlow.Id

#Defines arrays
$Callflows = @()
$CallHandlingAssociations = @()
$Schedules = @()

#Add data to Arrays
$Schedules += $afterHoursSchedule
$Callflows += $afterHoursCallFlow
$CallHandlingAssociations += $afterHoursCallHandlingAssociation

#Adds a holiday to the AutoAttendant your are creating for each holiday in $holidays.
foreach ($hol in $holidays)
{ 

    $title = $hol.title -replace " ","_"
    $variablename = $hol.title -replace " ","_"
    
    $StartDate = $hol.date | get-date -uformat %d/%m/%Y
    


$GreetingPrompt = New-CsOrganizationalAutoAttendantPrompt -TextToSpeechPrompt "Our offices are closed for Christmas from December 24 to December 26. Please call back later."
$MenuOption = New-CsOrganizationalAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic 
$Menu = New-CsOrganizationalAutoAttendantMenu -Name "$title Menu" -MenuOptions @($MenuOption)
$CallFlow = New-CsOrganizationalAutoAttendantCallFlow -Name "$title" -Greetings @($GreetingPrompt) -Menu $Menu

$dtr = New-CsOnlineDateTimeRange -Start "$StartDate"
$Schedule = New-CsOnlineSchedule -Name "$title" -FixedSchedule -DateTimeRanges @($dtr)

$CallHandlingAssociation = New-CsOrganizationalAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $Schedule.Id -CallFlowId $CallFlow.Id

    $Schedules += $Schedule
    $Callflows += $CallFlow
    $CallHandlingAssociations += $CallHandlingAssociation
}


$o=New-CsOrganizationalAutoAttendant -Name "ScriptTest" -LineUris @($lineUri) -DefaultCallFlow $defaultCallFlow -EnableVoiceResponse -Schedules $Schedules -CallFlows $Callflows -CallHandlingAssociations $CallHandlingAssociations -Language "en-US" -TimeZoneId "UTC" -Operator $operatorEntity
#Comment out til here after first run.

#Comment out the code bellow on the first run.
#Uppdates/creates new holiday for each holiday in $holidays.
foreach ($hol in $holidays)
{ 
    
    $title = $hol.title -replace " ","_"
    $variablename = $hol.title -replace " ","_"
    
    $StartDate = $hol.date | get-date -uformat %d/%m/%Y
  
    
$oaa = Get-CsOrganizationalAutoAttendant -PrimaryUri "sip:oaa_23423@contoso.com"

$GreetingPrompt = New-CsOrganizationalAutoAttendantPrompt -TextToSpeechPrompt "Our offices are closed for Christmas from December 24 to December 26. Please call back later."
$MenuOption = New-CsOrganizationalAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic 
$Menu = New-CsOrganizationalAutoAttendantMenu -Name "$title Menu" -MenuOptions @($MenuOption)
$CallFlow = New-CsOrganizationalAutoAttendantCallFlow -Name "$title" -Greetings @($GreetingPrompt) -Menu $Menu

$dtr = New-CsOnlineDateTimeRange -Start "$StartDate"
$Schedule = New-CsOnlineSchedule -Name "$title" -FixedSchedule -DateTimeRanges @($dtr)

$CallHandlingAssociation = New-CsOrganizationalAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $Schedule.Id -CallFlowId $CallFlow.Id

$oaa.CallFlows = $oaa.CallFlows + @($CallFlow)
$oaa.Schedules = $oaa.Schedules + @($Schedule)
$oaa.CallHandlingAssociations = $oaa.CallHandlingAssociations + @($CallHandlingAssociation)

Set-CsOrganizationalAutoAttendant -Instance $oaa
}

