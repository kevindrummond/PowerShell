[CmdletBinding()]
Param(
  [Parameter(Mandatory=$True,Position=1)]
  [String]$EmailAddress,
  [Parameter(Mandatory=$false)]
  [string]$Subject,
  [string]$BeforeDate,
  [string]$AfterDate,
  [switch]$AllItemsInRange,
  [switch]$SpecifyEntries,
  [switch]$ShowOrginizer,
  [string]$Organizer,
  [string]$Start,
  [string]$End,
  [switch]$ShowInvitees,
  [switch]$DeleteAppointments,
  [switch]$ShowAppDetails,
  [switch]$ShowFullItemDetails,
  [switch]$CreateReport
)

if ($CreateReport)
{
    $MailboxObject = get-mailbox $EmailAddress
    $MailboxDisplayname = $MailboxObject.Alias
}


#Import the EWS bits and bobs, connect to the mailbox and bind to the target folder
Import-Module "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
$Service.AutodiscoverUrl($EmailAddress,{$true})
$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress)
$service.HttpHeaders.Add("X-AnchorMailbox", $EmailAddress)
$RootFolderName = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$EmailAddress)
$Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$RootFolderName)
$Itemview = new-object Microsoft.Exchange.WebServices.Data.itemview(1000)

$FilterCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
if ($Subject)
{
    $SubjectFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject, $Subject)
    $FilterCollection.Add($SubjectFilter)
}
if ($BeforeDate)
{
    $Before = get-date $BeforeDate
    $itemFilterDateEnd = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeReceived, $Before)
    $FilterCollection.Add($itemFilterDateEnd)
}
else 
{
    $Before = get-date
}
If ($AfterDate)
{
    $After = get-date $AfterDate
    $itemFilterDateStart = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeReceived, $After)
    $FilterCollection.Add($itemFilterDateStart)
}
else 
{
    $After = (get-date).AddYears(-1)
}

if ($Start)
{
    $StartTime = get-date $Start
    $itemFilterStartTime = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.MeetingRequestSchema]::Start, $StartTime)
    $FilterCollection.Add($itemFilterStartTime)
}
else 
{
    $StartTime = (get-date).AddYears(-1)
}


if ($End)
{
    $EndTime = get-date $End
    $itemFilterEndTime = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.MeetingRequestSchema]::End, $EndTime)
    $FilterCollection.Add($itemFilterEndTime)
}
else 
{
    $EndTime = get-date
}

if ($Organizer)
{
    $OrganizerFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.MeetingRequestSchema]::Organizer, $Organizer)
    $FilterCollection.Add($OrganizerFilter)
}

if ($Subject -or $BeforeDate -or $AfterDate -or $Start -or $Organizer -or $End)
{
    $Items = $Folder.FindItems($FilterCollection,$ItemView)
}
else
{
    $Items = $Folder.FindItems($ItemView)
}

if ($AllItemsInRange)
{
    write-host ""
    #write-host "----------------------------------------------------------------------------" -fore white
    write-host "The following Items will be deleted:"
    #$Calview = new-object Microsoft.Exchange.WebServices.Data.calendarview($After,$Before,100)
    $Calview = new-object Microsoft.Exchange.WebServices.Data.calendarview($StartTime,$EndTime,1000)
    $CalItems = $Folder.FindAppointments($Calview)
    if ($Subject)
    {
        $CalItems = $CalItems | ?{$_.Subject -like "*$Subject*"}
    }

    if ($Organizer)
    {
        $CalItems = $CalItems | ?{$_.Organizer -like "*$Organizer*"}
    }
    
    $CalItems | select subject,start,end # | ?{$_.subject -like "*$Subject*"}
    if ($DeleteAppointments)
    {
        $Proceed = read-host "Proceed with deletions? (y/n)"
        if ($Proceed -eq "y")
        {
            foreach ($CalItem in $CalItems)
            {
                $CalItemSub = $CalItem.Subject
                $CalItemStart = $CalItem.Start
                $CalItemEnd = $CalItem.End
                write-host "Deleting Item $CalItemSub`t$CalItemStart`t$CalItemEnd" -fore yellow
                if ($ShowOrginizer)
                {
                    write-host "Organizer: " -fore yellow -NoNewline
                    write-host "$CalitemOrg" -fore White
                }
                $CalItem.delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
                if ($ShowInvitees)
                {
                    $CiTo = $CalItem.DisplayTo
                    $CiCC = $CalItem.DisplayCC
                    write-host "To: $CiTo" -fore Gray
                    write-host "CC: $CiCC" -fore Gray
                }
            }
        }
        else 
        {
            exit
        }
    }
}
elseif ($SpecifyEntries)
{
    #"Specify"
    $CICount = 0
    #$Calview = new-object Microsoft.Exchange.WebServices.Data.calendarview($After,$Before,100)
    $Calview = new-object Microsoft.Exchange.WebServices.Data.calendarview($StartTime,$EndTime,1000)
    $CalItems = $Folder.FindAppointments($Calview)
    if ($Subject)
    {
        $CalItems = $CalItems | ?{$_.Subject -like "*$Subject*"}
    }

    if ($Organizer)
    {
        $CalItems = $CalItems | ?{$_.Organizer -like "*$Organizer*"}
    }
    

    $CalItemArr = @()
    foreach ($CI in $CalItems)
    {
        $CalItemArr += $CI
    }
    write-host "No : Start Time`t`t`tEnd Time`t`tSubject" -fore white
    write-host "---------------------------------------------------------------------------------------" -fore white
    foreach ($CalItem in $CalItemArr)
    {
        if ($ShowFullItemDetails)
        {
            $CalItem | fl
        }
        $CalitemOrg = $CalItem.Organizer.name
        $CalItemSub = $CalItem.Subject
        $CalItemStart = $CalItem.Start
        $CalItemEnd = $CalItem.End
        $CalItemSens = $CalItem.Sensitivity
        $CalItemSFB = $CalItem.JoinOnlineMeetingUrl
        $CalItemtype = $CalItem.AppointmentType
        $CalItemResponseType = $CalItem.MyResponseType
        $CalItemReminder = $CalItem.IsReminderSet
        $CalItemLastModName = $CalItem.LastModifiedName


        write-host "$CICount  : $CalItemStart`t$CalItemEnd`t$CalItemSub" -fore yellow

        if($CreateReport)
        {
            $MailboxDisplayname
            $ReportDate = get-date -format "yyyy.MM.dd"
            add-content ".\$ReportDate-$MailboxDisplayname-CalenderItemReport.txt" "$MailboxDisplayname;$CalItemStart;$CalItemEnd;$CalItemSub"
        }

        if ($ShowOrginizer)
        {
            write-host "Organizer: " -fore yellow -NoNewline
            write-host "$CalitemOrg" -fore White
        }
        if ($ShowInvitees)
        {
            $CiTo = $CalItem.DisplayTo
            $CiCC = $CalItem.DisplayCC
            write-host "To: $CiTo" -fore Gray
            write-host "CC: $CiCC" -fore Gray
        }
        if ($ShowAppDetails)
        {
            write-host "Sensitivity: $CalItemSens"
            write-host "App Type: $CalItemtype"
            $CalItemResponseType
            $CalItemReminder
            $CalItemLastModName
            if ($CalItemSFB -ne $null)
            {
                write-host "SFB Link: $CalItemSFB" -fore Cyan
            }
        }
        $CICount++
    }
    write-host ""
    if ($DeleteAppointments -eq $True)
    {
        $ToDelete = Read-host "select a number to delete (default is 0)"
        if ($ToDelete -ge 0)
        {
            $SingleCIToDelete = $CalItemArr[$ToDelete]
            write-host "Deleting Item $CalItemStart`t$CalItemEnd`t$CalItemSub" -fore yellow
            $SingleCIToDelete.delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
        }
        else 
        {
            "Next time, pick a number"
            exit
        }
    }
}
else 
{
    $ItemArr = @()
    foreach ($I in $Items)
    {
        $ItemArr += $I
    }

    $ICount = 0
    write-host "No : Start Time`t`t`tEnd Time`t`tSubject`t`t`tItem Type" -fore white
    write-host "---------------------------------------------------------------------------------------" -fore white
    foreach ($SingleItem in $ItemArr)
    {
        #$SingleItem | fl
        if ($ShowFullItemDetails)
        {
            $SingleItem.load()
            $SingleItem | fl
            #$SingleItem.schema | fl # | get-member
            #$OMS = $SingleItem.OnlineMeetingSettings
            #$OMS | fl
        }
        $SIOrg = $SingleItem.Organizer
        $SISub = $SingleItem.Subject
        $SIStart = $SingleItem.start
        $SIEnd = $SingleItem.end
        $SIType = $SingleItem.AppointmentType
        write-host "$Icount  : $SIStart`t$SIEnd`t$SiSub`t`t$SIType" -fore yellow
        if ($ShowOrginizer)
        {
            write-host "Organizer: " -fore yellow -NoNewline
            write-host "$SIOrg" -fore White
        }
        if ($ShowInvitees)
        {
            $SiTo = $SingleItem.DisplayTo
            $SICC = $SingleItem.DisplayCC
            write-host "To: $SiTo" -fore Gray
            write-host "CC: $SICC" -fore Gray
        }
        $ICount++
    }
    write-host ""
    if ($DeleteAppointments -eq $True)
    {
        $ToDelete = Read-host "select a number to delete (default is 0)"
        
        if ($ToDelete -ge 0)
        {
            $Cancel = "This meeting has been cancelled by the email Administrator"
            write-host "Deleting the following:" -fore Yellow
            $ItemArr[$ToDelete] | select subject,start,end,AppointmentType | ft -AutoSize

            try 
            {
                $ItemArr[$ToDelete].cancelmeeting($cancel)
            }
            catch 
            {
                $LastError = $Error[0]
                $ErrMess = $LastError.Exception
                write-host "Error thrown:"
                $ErrMess
                "OK, Let's just delete the thing then..."
                $ItemArr[$ToDelete].delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
            }
        }
        else 
        {
            "Next time, pick a number"
            exit
        }
    }
}
