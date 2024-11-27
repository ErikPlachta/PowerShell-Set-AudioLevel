# Using PowerShell to Manage Windows Audio Volume Level

Manage Windows Audio Level with PowerShell - A quick reference guide for controlling the audio level within Windows OS programmatically.

<img src=".bin/soundboard-mute-mika-baumeister.jpg" alt="This photo was taken by Mika Baumeister, and published on Unsplashed.com" style="border-radius: 6px; " width="700" height="467">

## Overview

This script was created because I needed a way to control the audio level within Windows remotely and programmatically. Windows does not offer a class or module to easily control the audio level directly, so I had to get creative. I found a way to control the audio level using the `SendKeys` method from the `WScript.Shell` COM object. This method allows you to send keystrokes to the active window, which in this case is the Windows OS.

### I've Broken Down the Content into Two Sections

1. The first section is a quick reference guide to the individual commands you can use to control the volume level within Windows.
2. The second section is a custom function you can use to control the volume level within Windows.

## Prerequisites

1. Windows OS
2. PowerShell
3. Basic understanding of PowerShell
4. Ability to run PowerShell scripts on your machine

## How to Identify the Key Codes

1. Go to https://unicodelookup.com/#173 and you'll see the HEX value of `173` is `0xAD`
   Taken 07/19/2021 from https://unicodelookup.com/2. Then go to https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes and search for the Hex value 0xAD. You'll see this is the Hex value for Volume Mute Key.

## Individual Commands

In this section, I'll show you how to control the volume level within Windows using PowerShell. You can use these commands in your scripts or run them directly in the PowerShell console.

### Toggle Mute

```powershell
(new-object -com wscript.shell).SendKeys([char]173)
```

### Volume Up

```powershell
(new-object -com wscript.shell).SendKeys([char]174)
```

### Volume Down

```powershell
(new-object -com wscript.shell).SendKeys([char]175)
```

## Custom Function - `Set-AudioLevel`

In this section, I've included the function I created and use. I've included the full function with commentary and a single line declaration for easy CLI usage.

### Fully Documented Version

> This is the full function from my notes that I use to control the volume level within Windows. I've included it here with full commentary for easy reference.

```powershell
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
```

### Single Line Declaration

> This is the same function as above converted to a single line for easy CLI usage.

```powershell
function Set-AudioLevel {param([Alias("AudioLevel","L","l","volume","vol")][Parameter(Mandatory,Position=1)][System.Double]$Level) {try{$wshShell=New-Object -ComObject wscript.shell;1..50|ForEach-Object{$wshShell.SendKeys([char]174);Start-Sleep -Milliseconds 5};$upPresses=$level/2.0;for($i=0;$i -lt $upPresses;$i+=1){$wshShell.SendKeys([char]175);Start-Sleep -Milliseconds 5};"SUCCESS: Volume set to approximately $level%."}catch{Write-Output "ERROR: Unable to set volume level.";Write-Error $_}}}
```

---

## Wrapping Up

If you made it this far, thank you for reading. I hope you found this guide helpful. If you want to collaborate, discuss, or have questions just let me know in comments below or on my GitHub repo: https://github.com/ErikPlachta/PowerShell-Set-AudioLevel.
