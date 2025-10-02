$csvPath = "U:\General Portfolio\No_Match\matched_people.csv"
$rootPath = "U:\General Portfolio\No_Match"

Import-Csv -Path $csvPath | ForEach-Object {
    $matchVal = $_.MatchingValue
    $last = $_.LastName
    $first = $_.FirstName
    $personId = $_.PersonID   # not $pid (reserved)

    $oldFolder = Join-Path $rootPath $matchVal
    if (Test-Path $oldFolder) {
        $newName = "$matchVal - $last,$first - $personId"
        Rename-Item -Path $oldFolder -NewName $newName
        Write-Output "Renamed: $oldFolder -> $newName"
    } else {
        Write-Output "No folder found for $matchVal"
    }
}
