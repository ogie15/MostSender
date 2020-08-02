Clear-Host
$send = Get-MessageTrace -RecipientAddress "globaladmin@ojokougbo.onmicrosoft.com" -StartDate 07/23/2020 -EndDate 08/02/2020 | Sort-Object -Property 'SenderAddress'
 
$NumberOfEmails = 1

$CheckerTwo = " "

$PreviousSendersAddress = " "

$TotalCount = $send.Count

$Ncount = 0

$FinalArray = @()

foreach ($item in $send) {
    $Ncount +=1
    $CheckerOne = $item.SenderAddress
    if($CheckerOne -match $CheckerTwo){
        $NumberOfEmails +=1
    }elseif($CheckerOne -notmatch $CheckerTwo){
        if($PreviousSendersAddress -ne " "){
            # Write-Host $PreviousSendersAddress "has sent" $NumberOfEmails
            $FinalArray += [PSCustomObject][Ordered]@{
                "Recipient Address" = $item.RecipientAddress
                "Sender Address" = $PreviousSendersAddress
                "Number Of Emails" = $NumberOfEmails
            }
        }
        $PreviousSendersAddress = $item.SenderAddress
        $CheckerTwo = $item.SenderAddress
        $NumberOfEmails = 1
    }
    if($Ncount -eq $TotalCount) {
        if($PreviousSendersAddress -ne " "){
            # Write-Host $PreviousSendersAddress "has sent" $NumberOfEmails
            $FinalArray += [PSCustomObject][Ordered]@{
                "Recipient Address" = $item.RecipientAddress
                "Sender Address" = $PreviousSendersAddress
                "Number Of Emails" = $NumberOfEmails
            }
        }
    }
}
Write-Output $FinalArray | Sort-Object 'Number Of Emails' -Descending

Write-Host "Jesus is Lord"
