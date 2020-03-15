#original code from https://thescriptkeeper.wordpress.com/2013/09/27/import-a-bunch-of-ics-calendar-files-with-powershell/

#improvements by Chris Givens (@givenscj)

function Convert-UTCtoLocal
{
param(
[parameter(Mandatory=$true)]
[String] $UTCTime
)
    $strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
    $TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
    $LocalTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TZ)
    return $localtime;
}

function ParseValue($line, $startToken, $endToken)
{
    if ($startToken -eq $null)
    {
        return "";
    }

    if ($startToken -eq "")
    {
        return $line.substring(0, $line.indexof($endtoken));
    }
    else
    {
        try
        {
            $rtn = $line.substring($line.indexof($starttoken));
            return $rtn.substring($startToken.length, $rtn.indexof($endToken, $startToken.length) - $startToken.length).replace("`n","").replace("`t","");
        }
        catch [System.Exception]
        {
            $message = "Could not find $starttoken"
            #write-host $message -ForegroundColor Yellow
        }
    }

}


function ConvertObject($data)
{
    $str = "";
    foreach($c in $data)
    {
        $str += $c;
    }

    return $str;
}

# Put all your ICS files in one folder and set that here:
$ICSpath="C:\Users\given\Downloads\MVPSummit2020-ICS\MVPSummit2020-ICS"
$ICSlist = get-childitem $ICSPath

#set your filters...
$filter = @("AI", "AZR", "AZRC", "AZRWD", "BZAP", "CDM", "DATA", "DTEC", "EM", "OAS", "ODEV", "WDEV", "WI");
$filter = @("CDM");

$outlook = new-object -com Outlook.Application;
$calendar = $outlook.Session.GetDefaultFolder(9) 
    
Foreach ($i in $ICSlist ){
     $file= $i. fullname
     $data = @{}
     $content = Get-Content $file -Encoding UTF8
     $content |
     
    foreach-Object {
      if($_.Contains(':')){
            $z=@{ $_.split( ':')[0] =( $_.split( ':')[1]).Trim()}
           $data.Add($z.Keys,$z.Values)
       }
     }

     $Subject = ($data.getEnumerator() | ?{ $_.Name -eq "SUMMARY"}).Value
     $subject = convertobject $subject;

     $valid = $false;

     foreach($f in $filter)
    {
        if ($subject.contains($f))
        {
            $valid = $true;
        }
    }

    if (!$valid)
    {
        write-host "Skipping - $subject";
        continue;
    }

    write-host "Processing - $subject";

     foreach($c in $content)
     {
        if ($c.startswith("DESCRIPTION"))
        {
            $body = $c.replace("DESCRIPTION:","")
            $body = $body.replace("\r", "`r");
            $body = $body.replace("\n", "`n");
            $body = $body.replace("&#43;", "+");
        }

        if ($c.startswith("X-ALT-DESC;FMTTYPE=text/html:"))
        {
            $altbody = $c.replace("X-ALT-DESC;FMTTYPE=text/html:","")
            $altbody = $altbody.replace("\r", "");
            $altbody = $altbody.replace("\n", "");

            <#
            $doc = New-Object -com "HTMLFILE";
            $doc.IHTMLDocument2_write($altbody);
            $altbody = $doc.outerhtml;

            $altBody = $doc.documentElement.outerHTML;

            remove-item "altbody" -ea SilentlyContinue;
            add-content "altbody" $altbody;

            $altBody = get-content "altbody" -Raw;
            #>

            $teamsLink = ParseValue $altbody "https://teams.microsoft.com/l/meetup-join/" "`"";
            $teamslink = "https://teams.microsoft.com/l/meetup-join/" + $teamslink;
            
            $callInLink = ParseValue $altbody "https://dialin.teams.microsoft.com/" "`"";
            $callInLink = "https://dialin.teams.microsoft.com/" + $callInLink;

            $altBody = "Teams Link: "+ $teamsLink + "`r`r" + "CallIn Link: " + $callInLink + "`r`r";
        }
     }
     
    $appt = $calendar.Items.Add(1)
 
 <#
     # The body spacing/encoding was a PAIN, excuse the ugliness.
     $Body=[regex]::match($content,'(?<=\DESCRIPTION:).+(?=\DTEND:)', "singleline").value.trim()
     $Body= $Body -replace "\r\n\s"
     $Body = $Body.replace("\,",",").replace("\n"," ")
     $Body= $Body -replace "\s\s"
     #>
 
     $Start = ($data.getEnumerator() | ?{ $_.Name -eq "DTSTART"}).Value -replace "T" -replace "Z"
     $Start = [datetime]::ParseExact($Start ,"yyyyMMddHHmmss" ,$null )

     $start = Convert-UTCtoLocal $start;
 
     $End = ($data.getEnumerator() | ?{ $_.Name -eq "DTEND"}).Value -replace "T" -replace "Z"
     $End = [datetime]::ParseExact($End ,"yyyyMMddHHmmss" ,$null )

     $end = Convert-UTCtoLocal $end;


     $Location = ($data.getEnumerator() | ?{ $_.Name -eq "LOCATION"}).Value
 
     $appt.Start = $Start
     $appt.End = $End
     
     $appt.Subject = $Subject;
     $appt.Categories = "Presentations" #Pick your own category!
     $appt.BusyStatus = 0   # 0=Free
     $appt.Location = $Location
     $appt.BodyFormat = 1;
     $appt.Body = $altbody.trim() + $body;
     $appt.ReminderMinutesBeforeStart = 15 #Customize if you want 
 
    
    $appt.Save()

    if ($appt.Saved)
        { write-host "Appointment saved. - $($app.Subject)"}
    Else {write-host "Appointment NOT saved."}
}
    
   