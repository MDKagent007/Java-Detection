#********************************************************************************
#*                                                                              *
#* THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"  *.
#* AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE    *
#* IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE   * 
#* ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE     *
#* LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR          *
#* CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF         * 
#* SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS     *
#* INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN      *
#* CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)      *
#* ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE   *
#* POSSIBILITY OF SUCH DAMAGE.                                                  *
#*                                                                              *
#* THIS SCRIPT IS THE PROPERTY OF ERIC R JAGIELSKI AND IT IS NOT TO BE          *
#* DISTRIBUTED WITHOUT CONSENT OF THE ORIGINAL CREATOR.                         *
#*                                                                              *
#* Copyright 2023 by Eric R. Jagielski. All Rights Reserved.                    *
#********************************************************************************

$serverName = "<SERVER_NAME>"      #Replace <SERVER_NAME> with actual server name where the scripts are all located.
$networkShare = "<NETWORK_SHARE>"  #Replace <NETWORK_SHARE> with the actual server network share, hidden or not; with at least the service account running the script or all users having read/write permissions

#****************************************************************************************************************************************************************
# The following 3 lines illustrate the code that needs to be executed locally on the server or workstation that is being scanned for Java versions
#Set-ExecutionPolicy -ExecutionPolicy Unrestricted -force      #forces 'Unrestricted' policy on the local system temporarily in order to allow script execution
#powershell -executionpolicy bypass -file "\\" + $serverName + "\" + networkShare + "\JavaDetect.ps1"
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -force      #forces 'RemoteSigned' policy after script execution has completed
#****************************************************************************************************************************************************************

# Define the output files
$summaryFile = "\\" + $serverName + "\" + $networkShare + "\JavaVersionsSummary.txt"
$detailFile = "\\" + $serverName + "\" + $networkShare + "\Systems\" + $env:COMPUTERNAME + "_JavaDetails.txt"

# Function to format each line for the text file
Function Format-Line($javaDetail) {
    return "Path: $($javaDetail.Path)`r`nVersion: $($javaDetail.Version)`r`nProduct Name: $($javaDetail.ProductName)`r`nQuantity: $($javaDetail.Quantity)`r`n-----------------------`r`n"
}

# Function to read existing data from a file
Function Read-ExistingData($filePath) {
    $existingData = @{}
    if (Test-Path $filePath) {
        $currentDetail = @{}
        Get-Content $filePath | ForEach-Object {
            if ($_ -match '^Path: (.*)$') { 
                if ($currentDetail.Count -gt 0) {
                    $existingData[$currentDetail.Path] = $currentDetail
                    $currentDetail = @{}
                }
                $currentDetail.Path = $matches[1]
            } elseif ($_ -match '^Version: (.*)$') { $currentDetail.Version = $matches[1] }
            elseif ($_ -match '^Product Name: (.*)$') { $currentDetail.ProductName = $matches[1] }
            elseif ($_ -match '^Quantity: (\d+)$') { $currentDetail.Quantity = [int]$matches[1] }
        }
        if ($currentDetail.Count -gt 0) {
            $existingData[$currentDetail.Path] = $currentDetail
        }
    }
    return $existingData
}

# Read existing data from both files
$existingSummaryDetails = Read-ExistingData $summaryFile
$existingDetailDetails = Read-ExistingData $detailFile

# Create a hashtable to store new Java details
$javaDetails = @{}

# Find all java.exe files and process each one
Get-Childitem -Path 'C:\' -Filter 'java.exe' -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
    # Get the version and product name
    $fileInfo = $_ | Get-Item -ErrorAction SilentlyContinue
    $version = $fileInfo.VersionInfo.ProductVersion
    $productName = $fileInfo.VersionInfo.ProductName
    $path = $fileInfo.FullName
    
    # Check if this Java version with the same path is already in our hashtable
    if ($javaDetails.ContainsKey($path)) {
        # Increment quantity for existing item
        $javaDetails[$path].Quantity++
    } else {
        # Add new item to the hashtable
        $javaDetails[$path] = [PsCustomObject]@{
            'Path'        = $path
            'Version'     = $version
            'ProductName' = $productName
            'Quantity'    = 1
        }
    }
}

# Merge and update existing data with new data for both summary and detail files
foreach ($path in $javaDetails.Keys) {
    # For summary file
    if ($existingSummaryDetails.ContainsKey($path)) {
        $existingSummaryDetails[$path].Quantity += $javaDetails[$path].Quantity
    } else {
        $existingSummaryDetails[$path] = $javaDetails[$path]
    }

    # For detail file
    $existingDetailDetails[$path] = $javaDetails[$path]
}

# Write unique details back to the summary file
$existingSummaryDetails.Values | ForEach-Object { Format-Line $_ } | Out-File $summaryFile

# Write details to the detail file (this will overwrite with the latest details)
$existingDetailDetails.Values | ForEach-Object { Format-Line $_ } | Out-File $detailFile
