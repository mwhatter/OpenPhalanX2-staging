$comboBoxComputerName.Items.Clear()
    $script:allComputers = .\Tools\Scripts\small_dol_gethost.ps1
    $comboBoxComputerName.Items.AddRange($script:allComputers)

function Get-AllComputers {
    # Get a list of all domain in the forest
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