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

# Define the summary text file and the output Excel file
$summaryFile = "\\" + $serverName + "\" + $networkShare + "\JavaVersionsSummary.txt"
$outputExcelFile = "\\" + $serverName + "\" + $networkShare + "\JavaVersionsSummary.xlsx"

# Check if the ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module is not installed. Installing now..."
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

# Initialize an array to hold the data
$excelData = @()
$javaDetail = @{}

# Read the summary file and convert it into objects
Get-Content $summaryFile | ForEach-Object {
    if ($_ -match '^Path: (.*)$') {
        # If a new 'Path' is found and the current hashtable is not empty, add it to the array
        if ($javaDetail.Count -gt 0) {
            $excelData += New-Object PSObject -Property $javaDetail
            $javaDetail = @{}
        }
        $javaDetail['Path'] = $matches[1]
    } elseif ($_ -match '^Version: (.*)$') { 
        $javaDetail['Version'] = $matches[1]
    } elseif ($_ -match '^Product Name: (.*)$') { 
        $javaDetail['ProductName'] = $matches[1]
    } elseif ($_ -match '^Quantity: (\d+)$') { 
        $javaDetail['Quantity'] = [int]$matches[1]
    }
}

# Add the last detail entry to the array
if ($javaDetail.Count -gt 0) {
    $excelData += New-Object PSObject -Property $javaDetail
}

# Export the data to an Excel file
$excelData | Export-Excel -Path $outputExcelFile -AutoSize -TableName "JavaVersions" -Show
