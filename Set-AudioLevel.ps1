function Set-AudioLevel {
    <# 
    .SYNOPSIS
        Manage Windows Audio Volume Levels With PowerShell.

    .DESCRIPTION
        Using PowerShell to create a COM Object of the type Windows Shell. Then running Windows Shell function function SendKeys() with the parameters `[char]173`, `[char]174`, or `[char]175`.

    .PARAMETER Level
        The desired volume level as a percentage (0-100).

    .EXAMPLE
        # Set Audio Level to a target percentage
        Set-AudioLevel -Volume 60

        # Mute Audio
        Set-AudioLevel -Volume 0

    .NOTES 
        Author:     Erik Plachta
        Created:    11/09/2021
        Updated:    20241127
        Version:    0.0.22
        Changelog:
            - 0.0.1 | 20211109 | Erik Plachta | FEAT: Initial Version
            - 0.0.2 | 20241026 | Erik Plachta | BUG: Fix volume level calculation rounding error.
            - 0.0.21| 20241026 | Erik Plachta | FEAT: Add validation. Add updated logic.
            - 0.0.22| 20241127 | Erik Plachta | CHORE: Cleanup and verify for publication to medium.com and GitHub readme.
    #>
    param(
        [Alias("AudioLevel", "L", "l", "volume", "vol")]    # Allow multiple parameter names
        [Parameter(Mandatory,Position = 1)]                 # Make the parameter mandatory and positional so can be used without specifying the parameter name
        [System.Double]$Level                               # Define the parameter type
    )

    try{

        # 1. Validate input to ensure level is between 0 and 100
        if ($level -lt 0 -or $level -gt 100) {
            Write-Output "Error: Volume level must be between 0 and 100."
            return
        }

        # 2. Create Shell Object
        $wshShell = New-Object -ComObject wscript.shell

        # 3. Set volume to minimum (0%) by sending Volume Down key repeatedly
        1..50 | ForEach-Object {
            $wshShell.SendKeys([char]174)  # [char]174 is Volume Down
            Start-Sleep -Milliseconds 5   # Small delay to ensure each key press registers
        }

        # 4. Calculate the exact number of Volume Up presses needed
        $upPresses = $level / 2.0  # Calculate as a double to avoid rounding

        # 5. Increment to the desired volume level using exact decimal count
        for ($i = 0; $i -lt $upPresses; $i += 1) {  # Increment by 0.5 for more precision
            $wshShell.SendKeys([char]175)  # [char]175 is Volume Up
            Start-Sleep -Milliseconds 5   # Small delay to ensure each key press registers
        }
        Write-Output "SUCCESS: Volume set to approximately $level%."
    }
    catch {
        Write-Output "ERROR: Unable to set volume level."
        Write-Error $_
    }

}
