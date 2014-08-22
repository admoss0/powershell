
GetLocalDC

# Function to locate a DC in the local AD site
Function GetLocalDC {
	# Set $ErrorActionPreference to continue so we don't see errors for the connectivity test
	# $ErrorActionPreference = 'SilentlyContinue'
	
	# Get all the local domain controllers
	$LocalDCs = ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()).Servers
	
	# Create an array for the potential DCs we could use
	$PotentialDCs = @()
	
	# Check connectivity to each DC
	ForEach ($LocalDC in $LocalDCs) {
		# Create a new TcpClient object
		$TCPClient = New-Object System.Net.Sockets.TCPClient
		
		# Try connecting to port 389 on the DC
		$Connect = $TCPClient.BeginConnect($LocalDC.Name,389,$null,$null)
		
		# Wait 250ms for the connection
		$Wait = $Connect.AsyncWaitHandle.WaitOne(250,$False)                      

		# If the connection was succesful add this DC to the array and close the connection
		If ($TCPClient.Connected) {
			# Add the FQDN of the DC to the array
			$PotentialDCs += $LocalDC.Name

			# Close the TcpClient connection
			$Null = $TCPClient.Close()
		}
	}
	
	# Pick a random DC from the list of potentials
	$DC = $PotentialDCs #  | Get-Random
	
	# Return the DC
	Return $DC
}

