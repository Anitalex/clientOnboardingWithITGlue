Start-Transcript -Path "C:\onboard\qc_Transcript.rtf" -append -NoClobber
$verbosepreference = 'continue'
$dateString = Get-Date –format 'MM_dd_yyyy-HH_MM_ss'

#region UserInput
$client = Read-Host "What is the name of the client?"
$key = Read-Host "What is your API key foor IT Glue"
#endregion

#region ITGlue
$url = 'https://api.itglue.com'
$Headers = @{
    'x-api-key' = $key
    'Content-Type' = 'application/vnd.api+json'
}
#endregion

#region DefineOutput

$date = get-date
$qc_htmlfile = "C:\onboard\QC_$client-$dateString.html"

$qc_header = "
<html>
<head>
    <style type=`"text/css`">
    .good {color:green;}
    .bad {color:red;}
    .neutral {color:black;}
    </style>
    <title>$client Onboard QC</title>
</head>
<body>
<h2 align=center>Onboard Quality Control report for [$client] on $date</h2>
"
$qc_header | Out-File -FilePath $qc_htmlfile

#endregion

#region Functions

function New_Section {
    param($name,$column1header,$column2header)
    
    # this just starts a new section near the bottom for license keys as I can decode them
    $new_section = "
    <table align=center border=1 width=80%>
    <h2 align=center>$name</h2>
    <tr>
        <td><b><center>Quality Check Task</center></b></td>
        <td><b><center>Result (Green=GOOD, Red=BAD)</center></b></td>
        <td><b><center>Notes/Fix</center></b></td>
    "
    $new_section  | Out-File -FilePath $qc_htmlfile -append

}

function Get_Client {
    param($client)
    # get the organization
    $orgurl = "organizations?page[size]=100"
    $results = Invoke-RestMethod -Uri "$url/$orgurl" -Method 'GET' -Headers $headers
    $org = $results.data | where {$_.attributes.name -match "$client"}
    return $org
}

#endregion

#region Execution

######################################################################################
################################ Quick Notes #########################################

$org = Get_Client -client $client
# add quicknotes
$quicknotes = $org.attributes.'quick-notes'
$qn_output = "
    <table align=center border=1 width=80%>
    <h2 align=center>QuickNotes</h2>
    <tr>
        <td><b><center>$quicknotes</center></b></td>
    "
$qn_output  | Out-File -FilePath $qc_htmlfile -append

######################################################################################
################################ Locations ###########################################

# get the locations
New_Section -name 'Location Information'
$orgid = $org.id
$locationurl = "$url/organizations/$orgid/relationships/locations"
$locresults = Invoke-RestMethod -Uri $locationurl -Method 'GET' -Headers $headers
$ITGlocations = $locresults.data
$ITGlocations.attributes.name
$locationlist = $($ITGlocations.attributes.name) -join ', '

# list locations
if($ITGlocations.attributes.name -ne $null) {    
    $locationlist_output = "
        <tr>
            <td>List of locations</td>
            <td class=neutral>$locationlist</td>
            <td class=neutral>Correct</td>
        </tr>
        "
    } else {
    $locationlist_output = "
        <tr>
            <td>List of locations</td>
            <td class=bad>There are no locations</td>
            <td class=bad>Check ConnectWise too</td>
        </tr>
        "
    }
$locationlist_output | Out-File -FilePath $qc_htmlfile -append

# locations are synced with ConnectWise
$locationerrors = 0
foreach ($loca in $ITGlocations){
    $integrated = $loca.attributes.'psa-integration'
    if ($integrated -ne 'enabled'){
        $locationnotsynced_output = "
            <tr>
                <td>Locations not synced with ConnectWise</td>
                <td class=bad>$($loca.attributes.name) is not synced.  Status is $($loca.attributes.'psa-integration')</td>
                <td class=bad>Check ConnectWise too</td>
            </tr>
            "
        $locationnotsynced_output | Out-File -FilePath $qc_htmlfile -append
        $locationerrors += 1
    }
}
if ($locationerrors -eq 0){
    $locationnotsynced_output = "
            <tr>
                <td>Locations not synced with ConnectWise</td>
                <td class=good>All locations are synced!</td>
                <td class=good>Check ConnectWise too</td>
            </tr>
            "
     $locationnotsynced_output | Out-File -FilePath $qc_htmlfile -append
}


######################################################################################
################################ Sites ###############################################

# get the sites
New_Section -name 'Site Information'
$siteurl = "$url/flexible_asset_types?filter[name]=Sites"
$siteid = (Invoke-RestMethod -Uri $siteurl -Method 'GET' -Headers $headers).data.id

$clientsitesURL = "$url/flexible_assets?filter[flexible-asset-type-id]=$siteid&filter[organization-id]=$orgid&include=attachments"
$clientsites = Invoke-RestMethod -Uri $clientsitesURL -Method 'GET' -Headers $headers

# All locations have at least 1 site

foreach ($location in $ITGlocations){
    $found = 0
    foreach ($clientsite in $clientsites.data){
        if ($($clientsite.attributes.traits.location.values.name) -match $($location.attributes.name)){
        $found = 1
        }
    }
    if ($found -eq 1){
        $site_output = "
            <tr>
                <td>Location has a site</td>
                <td class=neutral>$($location.attributes.name) has a site</td>
                <td class=neutral>Correct</td>
            </tr>
            "
    } else {
        $site_output = "
            <tr>
                <td>Location has a site</td>
                <td class=bad>$($location.attributes.name) does NOT have a site</td>
                <td class=bad>Please check to make sure all sites exist</td>
            </tr>
            "
    }
    $site_output | Out-File -FilePath $qc_htmlfile -append
}

# all sites have hours of operations

foreach ($clientsite in $clientsites.data){
    
    if ($($clientsite.attributes.traits.'hours-of-operation') -eq $null){
        $HoO_output = "
            <tr>
                <td>Hours of Operations - $($clientsite.attributes.name)</td>
                <td class=bad>$($clientsite.attributes.name) does NOT have hours of operations set</td>
                <td class=bad>Please add the Hours of Operations</td>
            </tr>
            "
    } else {
        $HoO_output = "
            <tr>
                <td>Hours of Operations - $($clientsite.attributes.name)</td>
                <td class=good>$($clientsite.attributes.name) has hours of operations set</td>
                <td class=good>Correct!</td>
            </tr>
            "
    }
    $HoO_output | Out-File -FilePath $qc_htmlfile -append
}

# all sites have site contacts

foreach ($clientsite in $clientsites.data){
    
    if ($($clientsite.attributes.traits.'site-contact') -eq $null){
        $siteC_output = "
            <tr>
                <td>Site Contact - $($clientsite.attributes.name)</td>
                <td class=bad>$($clientsite.attributes.name) does NOT have a site contact set</td>
                <td class=bad>Please add the contact for the site</td>
            </tr>
            "
    } else {
        $siteC_output = "
            <tr>
                <td>Site Contact - $($clientsite.attributes.name)</td>
                <td class=good>$($clientsite.attributes.name) has a site contact</td>
                <td class=good>Correct!</td>
            </tr>
            "
    }
    $siteC_output | Out-File -FilePath $qc_htmlfile -append
}


# all sites have after hours access

foreach ($clientsite in $clientsites.data){
    
    if ($($clientsite.attributes.traits.'after-hours-access') -eq $null){
        $afterHours_output = "
            <tr>
                <td>After Hours Access - $($clientsite.attributes.name)</td>
                <td class=bad>$($clientsite.attributes.name) does NOT have information for after hours</td>
                <td class=bad>Please add after hours if necessary</td>
            </tr>
            "
    } else {
        $afterHours_output = "
            <tr>
                <td>After Hours Access - $($clientsite.attributes.name)</td>
                <td class=good>$($clientsite.attributes.name) has information on after hours access</td>
                <td class=good>Correct!</td>
            </tr>
            "
    }
    $afterHours_output | Out-File -FilePath $qc_htmlfile -append
}


# all sites have network diagrams

foreach ($clientsite in $clientsites.data){
    
    if ($($clientsite.attributes.traits.'network-diagrams') -eq $null){
        $afterHours_output = "
            <tr>
                <td>Network Diagram - $($clientsite.attributes.name)</td>
                <td class=bad>$($clientsite.attributes.name) does NOT have a network diagram</td>
                <td class=bad>Please add a network diagram</td>
            </tr>
            "
    } else {
        $afterHours_output = "
            <tr>
                <td>Network Diagram - $($clientsite.attributes.name)</td>
                <td class=good>$($clientsite.attributes.name) has a network diagram</td>
                <td class=good>Correct!</td>
            </tr>
            "
    }
    $afterHours_output | Out-File -FilePath $qc_htmlfile -append
}

# all sites have floorplans

foreach ($clientsite in $clientsites.data){
    
    if ($($clientsite.attributes.traits.'floorplans') -eq $null){
        $floorplans_output = "
            <tr>
                <td>Floorplan - $($clientsite.attributes.name)</td>
                <td class=bad>$($clientsite.attributes.name) does NOT have a floorplan</td>
                <td class=bad>Please add a floorplan</td>
            </tr>
            "
    } else {
        $floorplans_output = "
            <tr>
                <td>Floorplan - $($clientsite.attributes.name)</td>
                <td class=good>$($clientsite.attributes.name) has a floorplan</td>
                <td class=good>Correct!</td>
            </tr>
            "
    }
    $floorplans_output | Out-File -FilePath $qc_htmlfile -append
}

# all sites have key card

foreach ($clientsite in $clientsites.data){
    
    if ($($clientsite.attributes.traits.'do-we-have-a-key-or-access-card') -eq $null){
        $keycard_output = "
            <tr>
                <td>Key card - $($clientsite.attributes.name)</td>
                <td class=bad>$($clientsite.attributes.name) does NOT have a key card</td>
                <td class=bad>Please add a key card</td>
            </tr>
            "
    } else {
        $keycard_output = "
            <tr>
                <td>Key card - $($clientsite.attributes.name)</td>
                <td class=good>$($clientsite.attributes.name) has a key card</td>
                <td class=good>Correct!</td>
            </tr>
            "
    }
    $keycard_output | Out-File -FilePath $qc_htmlfile -append
}

# all sites have parking instructions

foreach ($clientsite in $clientsites.data){
    
    if ($($clientsite.attributes.traits.'parking-instructions') -eq $null){
        $keycard_output = "
            <tr>
                <td>Parking instructions - $($clientsite.attributes.name)</td>
                <td class=bad>$($clientsite.attributes.name) does NOT have parking instructions</td>
                <td class=bad>Please add parking instructions</td>
            </tr>
            "
    } else {
        $keycard_output = "
            <tr>
                <td>Parking instructions - $($clientsite.attributes.name)</td>
                <td class=good>$($clientsite.attributes.name) has parking instructions</td>
                <td class=good>Correct!</td>
            </tr>
            "
    }
    $keycard_output | Out-File -FilePath $qc_htmlfile -append
}


######################################################################################
################################ Facilities ###############################################

New_Section -name 'Facility Information'

$facilityurl = "$url/flexible_asset_types?filter[name]=Facility"
$facilityid = (Invoke-RestMethod -Uri $facilityurl -Method 'GET' -Headers $headers).data.id

$clientfacilityURL = "$url/flexible_assets?filter[flexible-asset-type-id]=$facilityid&filter[organization-id]=$orgid&include=attachments"
$clientfacilities = Invoke-RestMethod -Uri $clientfacilityURL -Method 'GET' -Headers $headers


foreach ($clientsite in $clientsites.data){
    $found = 0
    foreach ($clientfacility in $clientfacilities.data){
        if ($($clientfacility.attributes.traits.site.values.name) -like $($clientsite.attributes.name)){
        $found = 1
        }
    }
    if ($found -ne 0){
        $facility_output = "
            <tr>
                <td>Site has a facility</td>
                <td class=good>$($clientsite.attributes.name) has a facility</td>
                <td class=good>Correct</td>
            </tr>
            "
    } else {
        $facility_output = "
            <tr>
                <td>Site has a facility</td>
                <td class=bad>$($clientsite.attributes.name) does NOT have a facility</td>
                <td class=bad>Please check to make sure all sites have a facility</td>
            </tr>
            "
    }
    $facility_output | Out-File -FilePath $qc_htmlfile -append

}

foreach ($facility in $clientfacilities.data.attributes){
    $Name = $facility.name
    $equipment = $facility.traits.equipment.values
    if ($equipment -eq $null){
        $equipment_output = "
            <tr>
                <td>Facility Attached Equipment - $Name</td>
                <td class=bad>No equipment attached!</td>
                <td class=bad>Please attach the configurations for the equipment in the facility</td>
            </tr>
            "
    } else {
        $equipment_output = "
            <tr>
                <td>Facility Attached Equipment - $Name</td>
                <td class=good>$($equipment.name) </td>
                <td class=good>Correct!</td>
            </tr>
            "
    }
    $equipment_output | Out-File -FilePath $qc_htmlfile -append
}


#region OpenOutput

Write-Verbose (("Showing the output html at ") + (get-date))
Invoke-Item $qc_htmlfile

#endregion

#region Cleanup

Stop-Transcript
Get-Process | Where {$_.processname -match "cmd"} | Stop-Process

#endregion










