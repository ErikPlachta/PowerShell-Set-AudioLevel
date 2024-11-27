<# 
.SYNPOSIS
Manage Windows Audio Volume Levels With PowerShells

.DESCRIPTION
Using PowerShell to create a COM Object of the type Windows Shell. Then running Windows Shell function function SendKeys() with the parameters `[char]173`, `[char]174`, or `[char]175`.

.EXAMPLE Incrementing by Percentage       
    #Sets volume to 60%
    set_AudioLevel -Volume 60

    #Sets volume to 80%
    set_AudioLevel -Volume 80

    #Sets volume to 100%
    set_AudioLevel -Volume 100

.NOTES 
    Author:     Erik Plachta
    Created:    11/09/2021
    Updated:    10/26/2024
    Version:    0.0.2
    0.0.1 | 11/09/2021 | Erik Plachta | Initial Version
    0.0.2 | 10/26/2024 | Erik Plachta | FEAT: Add validation. Add updated logic. BUG: Fix volume level calculation.
#>
function Set-AudioLevel {
    param(
        [Alias("AudioLevel")][Parameter(Mandatory,Position = 1)][System.Double]$Level
    )

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
    $upPresses = $level / 4.0  # Calculate as a double to avoid rounding

    # 5. Increment to the desired volume level using exact decimal count
    for ($i = 0; $i -lt $upPresses; $i += 0.5) {  # Increment by 0.5 for more precision
        $wshShell.SendKeys([char]175)  # [char]175 is Volume Up
        Start-Sleep -Milliseconds 5   # Small delay to ensure each key press registers
    }

    Write-Output "Volume set to approximately $level%."

}


Set-AudioLevel -Level 100
