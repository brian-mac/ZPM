function ValidFileDate ($TargetFile,$DaysOld,$Sender,$Recipient)
{
    # Checks to see if the target file is older than a certian number of days, if so it will send a mail notifying a target recipient.
    $FileDate = (get-itemproperty -path ($TargetFile)).CreationTime
    $TodaysDate = get-date 
    
    if ($FileDate -lt $TodaysDate.AddDays(-$daysOld))
    {
        SendMail $Sender $Recipient "File is older than $($daysOld) days" "File was last updated on $($FileDate)"
    }

}
Function SendMail ($Sender, $target, $subject, $Body)
{
    $UnpackTarget =  (($Target.split("@").item(0)).replace("."," ")) + " <$Target>"
    send-mailmessage -from $Sender -to $UnpackTarget -subject $subject -Body $Body
}

# Check file to see if it is older than a certian date, requiers full path to the file, then the number of day old it is allowed to be, mail sender and receiver.

ValidFileDate "C:\temp\DepartmentsFN1.csv" 7 "Chatter Server <SYDW@star.com.au>" "Brian.mcelhinney@star.com.au"
