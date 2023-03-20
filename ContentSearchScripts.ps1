#Create Content Case
New-ComplianceCase -Name "Arraya Request #950964"

#Create Content Search 
##Content Search for Mailboxes
 ### Content Search for Mailboxes in a Group

#Create Content Search for Search Terms
## Check if more than 500 mailboxes in search
function  New-ComplianceContentCaseSearch {
    param (
        [Parameter(Mandatory=$True,HelpMessage="Import-CSV File Path")] [string] $ImportCSV
    )
    # Do a quick check to make sure our group name will not collide with other searches
    $searchCounter = 1
    $allSearchTermImport = import-csv $ImportCSV
    foreach ($searchterm in $allSearchTermImport) {
        $searchName = $searchterm.case +'_' + $searchCounter
        #Search Check
            if ($searchCheck = Get-ComplianceSearch $searchName -EA silentlycontinue) {
                Write-Warning "The Search Group $($searchName) conflicts with existing searches. Skipping"
            }
            else {
                # Create the query
                $query = $searchterm.ContentMatchQuery
                if(($searchterm.StartDate -or $searchterm.EndDate))  {
                    # Add the appropriate date restrictions.  NOTE: Using the Date condition property here because it works across Exchange, SharePoint, and OneDrive for Business.
                    # For Exchange, the Date condition property maps to the Sent and Received dates; for SharePoint and OneDrive for Business, it maps to Created and Modified dates.
                    if($query)
                    {
                        $query += " AND"
                    }
                    $query += " ("
                    if($searchterm.StartDate)
                    {
                        $query += "Date>=" + $searchterm.StartDate
                    }
                    if($searchterm.EndDate)
                    {
                        if($searchterm.StartDate)
                        {
                            $query += " AND "
                        }
                        $query += "Date<=" + $searchterm.EndDate
                    }
                    $query += ")"
                }
                # -ExchangeLocation can't be set to an empty string, set to null if there's no location.
                $exchangeLocation = $null
                if ($searchterm.ExchangeLocation) {
                        $exchangeLocation = $searchterm.ExchangeLocation
                }
                else {
                    $exchangeLocation = "All"
                }
                # Create and run the search
                Write-Host "Creating and running search: " $searchName -NoNewline
                $search = New-ComplianceSearch -Case $searchterm.Case -Name $searchName -ExchangeLocation $exchangeLocation -ContentMatchQuery $query -ea stop

                # Start and wait for each search to complete
                Start-ComplianceSearch $search.Name
                while ((Get-ComplianceSearch $search.Name).Status -ne "Completed")
                {
                    Write-Host " ." -NoNewline
                    Start-Sleep -s 3
                }
            }
    Write-Host ""
    $searchCounter++
    }
}

function  New-ComplianceContentCaseSearch {
    param (
        [Parameter(Mandatory=$True,HelpMessage="Import-CSV File Path")] [string] $ImportCSV
    )
    # Do a quick check to make sure our group name will not collide with other searches
    $allSearchTermImport = import-csv $ImportCSV
    foreach ($searchterm in $allSearchTermImport) {
        $searchName = $searchterm.ExchangeLocation
        #Search Check
            if ($searchCheck = Get-ComplianceSearch $searchName -EA silentlycontinue) {
                Write-Warning "The Search Group $($searchName) conflicts with existing searches. Skipping"
            }
            else {
                # Create the query
                $query = $searchterm.ContentMatchQuery
                $exchangeLocation = $searchterm.ExchangeLocation

                # Create and run the search
                Write-Host "Creating and running search: " $searchName -NoNewline
                $search = New-ComplianceSearch -Case $searchterm.Case -Name $searchName -ExchangeLocation $exchangeLocation -ContentMatchQuery $query -ea stop

                # Start and wait for each search to complete
                Start-ComplianceSearch $search.Name
                while ((Get-ComplianceSearch $search.Name).Status -ne "Completed")
                {
                    Write-Host " ." -NoNewline
                    Start-Sleep -s 3
                }
            }
    Write-Host ""
    }
}

#Large Freedom Mortgage Search
function Get-ComplianceSearchMaxResultsCheck {
    Param(
        [Parameter(Mandatory=$false,HelpMessage="Specify ContentSearchName")] [string] $searchCase,
        [Parameter(Mandatory=$false,HelpMessage="Specify ContentSearchName")] [string] $searchName,
        [Parameter(Mandatory=$True,HelpMessage="Import-CSV File Path")] [string] $ImportCSV
   )
   if ($searchCase) {
    $searches = Get-ComplianceSearch -Case $searchCase
   }
   elseif ($searchName) {
    $searches = Get-ComplianceSearch $searchName
   }
   else {
    $searches = Get-ComplianceSearch
   }
   foreach ($search in $searches)
   {
    if ($search.Status -ne "Completed")
    {
                    "Please wait until the search finishes.";
                    break;
    }
    $results = $search.SuccessResults;
    if (($search.Items -le 0) -or ([string]::IsNullOrWhiteSpace($results)))
    {
                    "The compliance search " + $search.name + " didn't return any useful results.";
                    break;
    }
    $mailboxes = @();
    $lines = $results -split '[\r\n]+';
    foreach ($line in $lines)
    {
        if ($line -match 'Location: (\S+),.+Item count: (\d+)' -and $matches[2] -gt 0)
        {
            $mailboxes += $matches[1];
        }
    }
    "Number of mailboxes that have search hits: " + $mailboxes.Count
    }
   
}

#Get Search Stats
function Get-ComplianceSearchStatistics {
    param (
        [Parameter(Mandatory=$false,HelpMessage="Specify ContentSearchName")] [string] $searchGroup,
        [Parameter(Mandatory=$True,HelpMessage="Specify ContentSearch Case")] [string] $searchCase,
        [Parameter(Mandatory=$false,HelpMessage="Import-CSV File Path")] [string] $ImportCSV,
        [Parameter(Mandatory=$True,HelpMessage="Export CSV File Path")] [string] $outputFile
    )
    if ($searchCase) {
        $searches = Get-ComplianceSearch -Case $searchCase
    }
    elseif ($searchGroup) {
        $searches = Get-ComplianceSearch | ?{$_.Name -clike $searchGroup + "_*"}
    }
    
    $allSearchStats = @()
    foreach ($partialObj in $searches) {
        $search = Get-ComplianceSearch $partialObj.Name
        $sizeMB = [System.Math]::Round($search.Size / 1MB, 2)
        $sizeGB = [System.Math]::Round($search.Size / 1GB, 3)
        $searchStatus = $search.Status
        if($search.Errors)  {
            $searchStatus = "Failed"
        }
        elseif($search.NumFailedSources -gt 0) {
            $searchStatus = "Failed Sources"
        }
        $searchStats = New-Object PSObject
        Add-Member -InputObject $searchStats -MemberType NoteProperty -Name Name -Value $search.Name
        Add-Member -InputObject $searchStats -MemberType NoteProperty -Name ContentMatchQuery -Value $search.ContentMatchQuery
        Add-Member -InputObject $searchStats -MemberType NoteProperty -Name Status -Value $searchStatus
        Add-Member -InputObject $searchStats -MemberType NoteProperty -Name Items -Value $search.Items
        Add-Member -InputObject $searchStats -MemberType NoteProperty -Name "Size(MB)" -Value $sizeMB
        Add-Member -InputObject $searchStats -MemberType NoteProperty -Name "Size(GB)" -Value $sizeGB
        $allSearchStats += $searchStats
    }
    # Save the results to a CSV file
    if ($outputFile)
    {
    $allSearchStatsPrime | Export-Csv -Path $outputFile -NoTypeInformation
    }
}

function  New-ComplianceSearchContentSearch {
    param (
        [Parameter(Mandatory=$True,HelpMessage="Import-CSV File Path")] [string] $ImportCSV
    )
    # Do a quick check to make sure our group name will not collide with other searches
    $searchCounter = 1
    $allSearchTermImport = import-csv $ImportCSV
    foreach ($searchterm in $allSearchTermImport) {
        $searchName = $searchterm.case +'_' + $searchCounter
        #Search Check
            if ($searchCheck = Get-ComplianceSearch $searchName -EA silentlycontinue) {
                Write-Warning "The Search Group $($searchName) conflicts with existing searches. Skipping"
            }
            else {
                # Create the query
                $queries = 
                $query = $searchterm.ContentMatchQuery
                if(($searchterm.StartDate -or $searchterm.EndDate))  {
                    # Add the appropriate date restrictions.  NOTE: Using the Date condition property here because it works across Exchange, SharePoint, and OneDrive for Business.
                    # For Exchange, the Date condition property maps to the Sent and Received dates; for SharePoint and OneDrive for Business, it maps to Created and Modified dates.
                    if($query)
                    {
                        $query += " AND"
                    }
                    $query += " ("
                    if($searchterm.StartDate)
                    {
                        $query += "Date>=" + $searchterm.StartDate
                    }
                    if($searchterm.EndDate)
                    {
                        if($searchterm.StartDate)
                        {
                            $query += " AND "
                        }
                        $query += "Date<=" + $searchterm.EndDate
                    }
                    $query += ")"
                }
                # -ExchangeLocation can't be set to an empty string, set to null if there's no location.
                $exchangeLocation = $null
                if ($searchterm.ExchangeLocation) {
                        $exchangeLocation = $searchterm.ExchangeLocation
                }
                else {
                    $exchangeLocation = "All"
                }
                # Create and run the search
                Write-Host "Creating and running search: " $searchName -NoNewline
                $search = New-ComplianceSearch -Case $searchterm.Case -Name $searchName -ExchangeLocation $exchangeLocation -ContentMatchQuery $query -ea stop

                # Start and wait for each search to complete
                Start-ComplianceSearch $search.Name
                while ((Get-ComplianceSearch $search.Name).Status -ne "Completed")
                {
                    Write-Host " ." -NoNewline
                    Start-Sleep -s 3
                }
            }
    Write-Host ""
    $searchCounter++
    }
}

$FullArraySearches = $null
foreach ($term in $kqlSearches) {
    $FullArraySearches += "("
    $FullArraySearches += $term."Kql Syntax"
    $FullArraySearches += ")"
    $FullArraySearches += " OR "
}

$KQLALLSearchTerms = (("art show") OR ("Bra tini") OR ("Country Club") OR ("fun night") OR ("gift bags") OR ("Golden Nugget") OR ("happy hour") OR ("havana night") OR ("Lounge & Learn") OR ("Lounge AND Learn") OR ("Lunch & Learn") OR ("Lunch AND Learn") OR ("monte carlo") OR ("Pink Tie" OR "PinkTie") OR ("trade show") OR ("Triple Play") OR (advertisement) OR (Affinity) OR ("agent" ONEAR(2) "mixer") OR (appetizers) OR (awards) OR (banner) OR (banner*) OR (baseball) OR (bowling) OR (Brochure) OR (Brunch) OR (charity) OR (cocktail) OR (Conference) OR (Dance) OR (dinner) OR (Event) OR (Flyer) OR (foseball) OR (Fundraiser) OR (Funsports) OR (Gala) OR (Giveaway) OR (Golf) OR ("Holiday" ONEAR(2) (Dinner OR Party)) OR (karaoke) OR (Logo) OR (lunch) OR (membership) OR (Networking) OR (Outing) OR (Party) OR (picnic) OR (Pindar) OR (poker) OR (Raffle) OR (Seminar) OR ("speaking" NEAR (class OR engagement OR venue OR forum OR panel OR opportunity OR time OR event)) OR (Sponsor*) OR (sports NEAR(2) bar) OR (table NEAR(2) (name OR sponsor)) OR (tournament) OR (venue) OR (wine) OR (CardTapp) OR ("list report*") OR ("geodata" OR ("geo" NEAR(2) "data")) OR (invitation) OR (Simplay) OR (gifts) OR ("cobranded marketing") OR ("open house") OR ("get together") OR (lottery) OR (rental NEAR(2) credit) OR (screening) OR (tenants) OR ("Gold Nugget") OR ("Mohegan Sun") OR ("Mohegan Rally") OR (DJ) OR (Funkmaster) OR (invoice) OR (coffee) OR ("digital camera") OR (ticket*) OR (drinks) OR ("gift cards") OR ("continuing education") OR ("realtor class") OR (Leyritz) OR (LeTip) OR (Belmont w/ race) OR ("Villa Lombardi") OR (appreciation NEAR (lunch OR luncheon OR dinner OR party)) OR ("pizza party") OR (Hamlet) OR ("thank you cards") OR ("thank-you cards") OR (escape NEAR (room OR game)) OR (Caesar*) OR ("realtor expenses") OR ("movie night") OR (celebrity) OR (desk NEAR (rental OR lease)) OR (food) OR (snacks) OR ("open house" NEAR(5) suppl*) OR (Citrus Club) OR (Bourbon Heat) OR ("Meet AND Greet") OR ("Meet & Greet") OR ("Mind-set" OR "Mind-set training") OR ('Night at the Races') OR (Holiday NEAR(2) Event) OR (DiNapoli) OR (ReMax Edge) OR (Piatto) OR (Dumbo) OR (Vanderbilt) OR (Taco Tuesday) OR (Rosalitas) OR (Foundation Certs) OR (Wicked Monk) OR (Woodland Lock Change) OR (Sangria) OR (Rab's Country Lanes) OR (Brivity) OR (Boom Town) OR (Catering by Amy) OR (Millie's of Staten Island) OR (El Aguila Dorada) OR (Italian Pork Store) OR (KTV) OR (Cuba Libre) OR (King Umberto) OR (Press 195) OR (Patrizias) OR (Mansion) OR (Calogero*) OR (Taqueria Maria & Ricardo) OR (Johnny*) OR (Capital Grille) OR (Spa 88) OR (Cue Bar) OR (Oyster Bar) OR (Jade Asian Bistro) OR (Rosen Shingle Creek) OR (Ruth* Chris) OR (Valentino*) OR (Wolfgang Puck) OR (El Pilon) OR (Madrona) OR (Estrella Latino) OR (La Isla) OR (Grand Wine Cellar) OR ("La Bottga" OR "La Bottega") OR (King Umberto) OR (One10) OR ("Morton's") OR ("Assaggio!") OR (Real Greeks) OR (Viranda) OR (La Piazza) OR (V I Pizza) OR ("Borrelli's Restaurant") OR ("Kwik Entertainment") OR ("Jake's steak") OR ("Mill Pond") OR (Pino's) OR ("Lobby Bar MGM") OR (Sabor peruano) OR (Chipper Truck Caf*) OR (Ragazzi Italian Kitchen) OR (Lillian Pizzeria) OR ("Thom Thom" OR "Thom Thom's") OR (Calabria) OR ("Marina View" OR "Marina di Calabria") OR ("Butcher Bar") OR ("La Focaccia") OR (Bulla) OR (Lillian Pizza) OR (Lillian Pizzeria) OR ("Nick & Stef's") OR (Kenzo) OR (Sushi) OR ("Vasili's") OR (Valley Caterers) OR (Olive Garden) OR (Fogo de Chao) OR (El Patio) OR (Portuguese Fisherman) OR (Salty Dog) OR (Mr Pollo) OR (Leonardos) OR (Sabor Peruano) OR (Toku) OR (Fushimi) OR (Yard House) OR (Aperitif) OR (Roslyn Social) OR (Parkville Plaza) OR (Parkville Deli) OR (Stagecoach Tavern) OR (Total Wine) OR (Top Golf) OR (Cory's Ale House) OR (369 Liquors) OR ("Hemingway's") OR (BJ's Brewhouse) OR (Meat Market Tampa) OR (Petes Brew House) OR (Petroleum Club) OR (Referral partners*) OR (Real Talk) OR (Rally) OR (Drawing) OR (Grand Opening) OR (Vineyard) OR (Vineyards) OR (Holiday gathering) AND (Date>=2017-01-01))