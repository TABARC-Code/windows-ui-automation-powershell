# windows-ui-automation-powershell
Pixel-based UI automation using native Win32 APIs. - Or a way to automate some AI Chat Functions )

# Windows UI Automation (PowerShell)

Pixel-based UI automation using native Win32 APIs.

This script:
1. Brings a target application window to the foreground
2. Clicks a fixed screen coordinate
3. Types text
4. Clicks a second coordinate
5. Repeats with a random delay

## Requirements

- Windows
- PowerShell 5.1 or PowerShell 7+
- Single, stable display setup recommended

## Configuration

Edit these values in the script:

```powershell
$TargetProcessName = "notepad"
$FirstClick  = @{ X = 353; Y = 981 }
$SecondClick = @{ X = 842; Y = 977 }
$TextToType  = "continue"

bsicaly use the x/y plotter to locate the sublit button location of the chatgtp ui and the location of the command system.

then just put thise details in the script replacng the exisiting cordinates. Then let the script just type next and click ok to run the chatgtp or what ever automatically. ts a simple dump script but gets a round a lot of issues.
