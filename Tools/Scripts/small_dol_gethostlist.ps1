#Aligned and ready for testing
# This script contains the logic to fetch all computers
function Get-AllComputers {
    # Get a list of all domains in the forest
    $domains = (Get-ADForest).Domains

    # Use ForEach-Object -Parallel to process the domains in parallel
    $domaincomputers = $domains | ForEach-Object -Parallel {
        # For each domain, get a DC
        $domainController = (Get-ADDomainController -DomainName $_ -Discover -Service PrimaryDC).HostName

        # Check if $domainController is a collection and get the first hostname
        if ($domainController -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
            $domainController = $domainController[0]
        }

        # Get enabled computers from the DC
        Get-ADComputer -Filter {Enabled -eq $true} -Server $domainController | Select-Object -ExpandProperty Name
    } | Sort-Object

    return $domaincomputers
}

# Call the function and return its output to be caught by the main script
Get-AllComputers
