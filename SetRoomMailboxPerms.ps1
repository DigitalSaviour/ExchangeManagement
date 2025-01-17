# Script to configure permissions and settings for a specified room mailbox
# Run this script in an environment where Exchange Management Shell or Exchange Online PowerShell module is available.

param (
    [Parameter(Mandatory = $true)]
    [string]$RoomMailbox
)

# Function to validate that the mailbox exists
function Validate-Mailbox {
    param (
        [string]$Mailbox
    )

    $mailboxExists = Get-Mailbox -Identity $Mailbox -ErrorAction SilentlyContinue
    if (-not $mailboxExists) {
        Write-Error "Mailbox '$Mailbox' does not exist. Please provide a valid room mailbox."
        exit 1
    }
}

# Function to configure room mailbox permissions
function Configure-RoomMailbox {
    param (
        [string]$Mailbox
    )

    try {
        Write-Host "Validating mailbox..." -ForegroundColor Cyan
        Validate-Mailbox -Mailbox $Mailbox

        Write-Host "Setting calendar processing to AutoAccept..." -ForegroundColor Cyan
        Set-CalendarProcessing -Identity $Mailbox -AutomateProcessing AutoAccept

        Write-Host "Removing ability for users to edit other people's bookings..." -ForegroundColor Cyan
        Set-MailboxFolderPermission -Identity "$Mailbox:\Calendar" -User Default -AccessRights Reviewer

        Write-Host "Configuration completed successfully for '$Mailbox'." -ForegroundColor Green
    } catch {
        Write-Error "An error occurred while configuring the room mailbox: $_"
    }
}

# Main Script Execution
Write-Host "Starting configuration for room mailbox: $RoomMailbox" -ForegroundColor Yellow
Configure-RoomMailbox -Mailbox $RoomMailbox
