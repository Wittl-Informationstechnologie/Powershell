# Azure Subnets
$ips = @("10.100.10.0")

# Point-To-Site IP address range

$azurePptpRange = "172.16.201."

# Find current IP address from Azure
$azureIpAddress = ipconfig | findstr $azurePptpRange

# If Azure hasn't given us one yet, exit and let u know
if (!$azureIpAddress){
    "You do not currently have an IP address in your Azure subnet."
    exit 1
}

$azureIpAddress = $azureIpAddress.Split(": ")
$azureIpAddress = $azureIpAddress[$azureIpAddress.Length-1]
$azureIpAddress = $azureIpAddress.Trim()

# Delete any previous configured routes for these ip ranges
foreach($ip in $ips) {
    $routeExists = route print | findstr $ip
    if($routeExists) {
        "Deleting route to Azure: " + $ip
        route delete $ip
    }
}

# Add new routes to Azure Virtual Network
foreach($subnet in $ips) {
    "Adding route to Azure: " + $subnet
    echo "route add $ip MASK 255.255.255.0 $azureIpAddress"
    route add $subnet MASK 255.255.255.0 $azureIpAddress
}