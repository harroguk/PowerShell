# This script will query a BGG Wishlist and grab prices for everything on it from boardgameatlas.
# It will then output whether the games are over/under the limit that is configured.

###########################
# SET THESE TWO VARIABLES #
###########################
# Limit in USD
$USDlimit=50
# User to be Queried
$BGGuser="jjz4gaming"
############################
# NOTHING TO DO BELOW HERE #
############################

# Get Wishlist
# Funky file handling is to deal with a UTF-8 error if you try to deal with the output in flight.
if(!(test-path "c:\temp")){
    mkdir c:\temp\ | out-null
}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
# Requests the wishlist. This can take some time before it becomes available for the first time.
do{
    invoke-restmethod "https://boardgamegeek.com/xmlapi2/collection?username=$BGGuser&wishlist=1" -outfile "c:\temp\wishlist.xml"
    [xml]$XmlDocument = cat "c:\temp\wishlist.xml"
    if($XmlDocument.message.ToString() -like "*Please try again later for access.*"){
		write-host "Waiting for Wishlist to become available, this can take some time."
        sleep 30
    }
}while($XmlDocument.message.ToString() -like "*Please try again later for access.*")
$wishlist = $XmlDocument.items.ChildNodes.name.innerxml

invoke-restmethod "https://boardgamegeek.com/xmlapi2/user?name=$BGGuser" -outfile "c:\temp\user.xml"
[xml]$XmlDocument = cat "c:\temp\user.xml"
$region = $XmlDocument.user.country.value

Function Get-ExchangeRate{
	Param(
		[string]$Currency
	)#End parameters
	Begin{
        $c = new-object system.net.WebClient
        [xml]$Lines = (New-Object system.io.StreamReader $c.OpenRead("https://www.bankofcanada.ca/valet/fx_rss/$Currency")).ReadToEnd()
        $rate = $lines.GetEnumerator().childnodes.statistics.exchangerate.value."#text"
		$rate
	}
}

$CADlimit = [int]$USDlimit*(Get-ExchangeRate "FXUSDCAD")

if($region -eq "Canada"){
    $outputObject = "price_ca"
    $limit = $CADlimit
}
elseif($region -eq "United Kingdom"){
    # set Limit in GBP
    $outputObject = "price_uk"
    $limit = [int]$CADlimit/(Get-ExchangeRate "FXGBPCAD")
}
elseif($region -eq "Australia"){
    # set Limit in AUD
    $outputObject = "price_au"
    $limit = [int]$CADlimit/(Get-ExchangeRate "FXAUDCAD")
}
else{
    $outputObject = "price"
    $limit = $USDlimit
}	

# Reset counters
$notavailable=0
$inlimit=0
$slightlyoverlimit=0
$overlimit=0
$Notread=0

write-host "Using Region           : " $region
Write-host "Limit in local currency: " ([int]$limit)

# Query Cost of games
$client_id="G497VGENgX"
foreach ($game in $wishlist){
    #sanitise names
    $InvertedValidCharactersRange = "[^A-Za-z0-9 ]"
    $game = $game -replace $InvertedValidCharactersRange, ""
	#query cost
	$cost = 999
	$cost = [int]@(((invoke-webrequest "https://api.boardgameatlas.com/api/search?client_id=$client_id&name=$game").content | ConvertFrom-json).games.$outputObject)[0]
	
	# Write Output
	write-host -nonewline "$game - "
	if($cost -eq 0){
	    write-host -foregroundcolor black -backgroundcolor red "Unavailable or Game not found"
		$notavailable++
	}
	elseif($cost -eq 999){
	    write-host -foregroundcolor black -backgroundcolor red "Error Reading Cost"
		$Notread++
	}
	elseif($cost -le $limit){
	    write-host -foregroundcolor black -backgroundcolor green $cost
		$inlimit++
	}
	elseif($cost -gt $limit -and $cost -lt $limit*1.5){
	    write-host -foregroundcolor black -backgroundcolor yellow $cost
		$slightlyoverlimit++
	}
	elseif($cost -ge $limit*1.5){
	    write-host -foregroundcolor black -backgroundcolor red $cost
		$overlimit++
	}
	
    if($wishlist.count -gt 40){
		# Slow Down the requests to avoid rate limiting.
	    sleep 1
	}
}

# Write Report
write-host
write-host "###########"
write-host "# Summary #"
write-host "###########"
write-host -foregroundcolor black -backgroundcolor red "Not Available: $notavailable"
write-host -foregroundcolor black -backgroundcolor green "Under Limit  : $inlimit"
write-host -foregroundcolor black -backgroundcolor yellow "Slightly Over: $slightlyoverlimit"
write-host -foregroundcolor black -backgroundcolor red "1.5x or more : $overlimit"
write-host -foregroundcolor black -backgroundcolor red "Errors       : $Notread"
write-host "Total        :"($notavailable + $inlimit + $slightlyoverlimit + $overlimit + $Notread)

pause
