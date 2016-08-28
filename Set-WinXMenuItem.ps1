Function Set-WinXMenuItem 
{
  <#
      .SYNOPSIS
      Add a shorcut to the Win-X menu

      .DESCRIPTION
      Creates a shorcut in $env:localappdata\Microsoft\Windows\WinX\[Group]\[Name] then hashlinks the .lnk and restarts explorer

      .PARAMETER Version
      Outputs the version of the script

      .PARAMETER Add
      Add a Win-X menu item

      .PARAMETER ApplicationPath
      The path of the application

      .PARAMETER Group
      The group to add the Win-X menu item for.

      .PARAMETER Name
      The name of the Win-X menu item.

      .PARAMETER Arguments
      Arguments to add to the shortcut

      .PARAMETER Remove
      Removes a Win-X menu item

      .PARAMETER ElevateShortcut
      Shortcut defaults to 'run as administrator'

      .PARAMETER HashLnkPath
      Path to the hashlnk tool if not in the script root.

      .EXAMPLE
      Set-WinXMenuItem -Version
      Output version of script

      .EXAMPLE
      Set-WinXMenuItem -Add -ApplicationPath "$env:windir\system32\SnippingTool.exe" -Group "Group3" -Name 'Snipping Tool'
      Add the snipping tool to the Win-X menu

      .EXAMPLE
      Set-WinXMenuItem -Add -ApplicationPath "C:\Tools\Utilities\Cmder\Cmder.exe" -Group "Group3" -Name 'Cmder' -ElevateShortcut $true
      Adds Cmder to the Win-X menu with the default option to run as administrator

      .EXAMPLE
      Set-WinXMenuItem -Remove -Name "Snipping Tool*"
      Remove 'Snipping Tool' from the Win-X menu

      .NOTES
      Author: Bevin Du Plessis
      Date: 28/08/2016
      Credits: Uses hashlnk tool by Rafael Rivera https://github.com/riverar/hashlnk

      .LINK
      https://github.com/nightshade2109/powershellscripts

      .INPUTS
      [switch]
      [string]
      [string]
      [string]
      [string]
      [bool]

      .OUTPUTS
      [string]
  #>


  [CmdletBinding(DefaultParameterSetName = 'AddOptions')]
  param 
  (
    [Parameter(ParameterSetName = 'Version')]
    [switch]$Version,
    
    [Parameter(ParameterSetName = 'AddOptions')]
    [switch]$Add,

    [Parameter(ParameterSetName = 'AddOptions',HelpMessage = 'Application path for shortcut',Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ApplicationPath,
    
    [Parameter(ParameterSetName = 'AddOptions',HelpMessage = 'Group for the shortcut',Mandatory = $true)]
    [ValidateSet('Group1', 'Group2', 'Group3', 'Group4', 'Group5', 'Group6', 'Group7', 'Group8', 'Group9', 'Group10')]
    [string]$Group,
    
    [Parameter(ParameterSetName = 'RemoveOptions',HelpMessage = 'Name of the shortcut',Mandatory = $true)]
    [Parameter(ParameterSetName = 'AddOptions',Mandatory = $true)]
    [string]$Name,
    
    [Parameter(ParameterSetName = 'AddOptions')]
    [string]$Arguments = $null,
    
    [Parameter(ParameterSetName = 'AddOptions')]
    [bool]$ElevateShortcut = $false,
    
    [Parameter(ParameterSetName = 'RemoveOptions')]
    [switch]$Remove,

    [Parameter(ParameterSetName = 'AddOptions')]
    [string]$HashLnkPath
  )
  
  

  Begin {
      
    Write-Verbose -Message "Starting $($MyInvocation.Mycommand)"

    $ScriptVersion = '0.1'

    if ($Version.IsPresent) 
    {
      $VersionResult = New-Object -TypeName PSObject -Property @{
        Version = $ScriptVersion
      }
      
      Write-Output -InputObject $VersionResult
    }
    
    Write-Verbose -Message $PSBoundParameters.Keys

    # Pass verbose debug to child functions
    if ($PSBoundParameters.ContainsKey('Debug'))
    {
      $Script:DebugPreference = 'Continue'
    }
    else 
    {
      $Script:DebugPreference = 'SilentlyContinue'
    }

    if ($PSBoundParameters.ContainsKey('Verbose'))
    {
      $Script:VerbosePreference = 'Continue'
    }
    else 
    {
      $Script:VerbosePreference = 'SilentlyContinue'
    }

  }

  Process 
  {
    Write-Verbose -Message "Processing $($MyInvocation.Mycommand)"
    
    if(!($Version.IsPresent)) 
    {
      if($Add.IsPresent) 
      {
        if([string]::IsNullOrEmpty($HashLnkPath)) 
        {
          Write-Verbose -Message 'hashlnk.exe path not specified attempting to find it in the script directory' 
          $HashLnkPath = Join-Path -Path $PSScriptRoot -ChildPath 'hashlnk.exe'
          Write-Verbose -Message $HashLnkPath
          
          if(Test-Path -Path $HashLnkPath)
          {
            Write-Verbose -Message "Found hashlnk.exe '$HashLnkPath'"
          } 
          else 
          {
            Write-Warning -Message 'hashlnk.exe not found'
            Write-Warning -Message 'Please specify location using the -HashLnkPath parameter'
            break
          }
        } 
        else 
        {
          if(Test-Path -Path $HashLnkPath)
          {
            Write-Verbose -Message "hashlnk.exe path is valid '$HashLnkPath'"
          } 
          else 
          {
            Write-Warning -Message 'hashlnk.exe not found'
            Write-Warning -Message 'Please specify location using the -HashLnkPath parameter'
            break
          }
        }
      }
    
      $GroupPath = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Microsoft\Windows\WinX'
      
      if($Add.IsPresent) 
      {
        $GroupPathName = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Microsoft\Windows\WinX\$Group"

        if(!([string]::IsNullOrEmpty($Arguments))) 
        {
          Write-Verbose -Message "Adding Win-X menu item '$Name' path '$ApplicationPath' with arguments '$Arguments' for group '$Group'"
        }
        else 
        {
          Write-Verbose -Message "Adding Win-X menu item '$Name' path '$ApplicationPath' for group '$Group'"
        }

        if(!(Test-Path -Path $GroupPathName)) 
        {
          Write-Verbose -Message "Path for '$Group' does not exist attempting to create '$GroupPathName'"
          try 
          {
            $null = New-Item -Path $GroupPathName -ItemType Directory -Force
          }
          catch 
          {
            Write-Verbose -Message "Unable to create '$Group' at path '$GroupPath'"
            break
          }
          finally 
          {
            Write-Verbose -Message "Path created '$GroupPathName'"
          }
        }
      
          
        if(Test-Path -Path $ApplicationPath) 
        {
          try 
          {
            if($ElevateShortcut) 
            {
              $ShortcutPath = Join-Path -Path "$GroupPathName" -ChildPath "$Name (Admin).lnk"
            }
            else 
            {
              $ShortcutPath = Join-Path -Path "$GroupPathName" -ChildPath "$Name.lnk"
            }
            
            Write-Verbose -Message "Creating lnk shotcut at '$ShortcutPath' for application '$ApplicationPath'"
            $WorkingDirectory = Split-Path -Path $ApplicationPath -Parent
            $WshShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut("$ShortcutPath")
            $Shortcut.TargetPath = "$ApplicationPath"
            $Shortcut.Arguments = "$Arguments"
            $Shortcut.IconLocation = "$ApplicationPath"
            $Shortcut.Description = $Name
            $Shortcut.WorkingDirectory = "$WorkingDirectory"
            $Shortcut.Save()

            if($ElevateShortcut) 
            {
              $bytes = [IO.File]::ReadAllBytes("$ShortcutPath")
              $bytes[0x15] = $bytes[0x15] -bor 0x20 #set byte 21 (0x15) bit 6 (0x20) ON
              [IO.File]::WriteAllBytes("$ShortcutPath", $bytes)
            }
          }
          catch 
          {
            Write-Warning -Message 'Unable to create shortcut'
            Write-Warning -Message "Error was $_"
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Warning -Message "Error was in Line $line"
            break
          }
          finally 
          {
            Write-Verbose 'Shortcut created'
          }
          
          try 
          {
            Write-Verbose -Message "Hashlinking '$ShortcutPath'"

            $pinfo = New-Object -TypeName System.Diagnostics.ProcessStartInfo
            $pinfo.FileName = $HashLnkPath
            $pinfo.RedirectStandardError = $true
            $pinfo.RedirectStandardOutput = $true
            $pinfo.UseShellExecute = $false
            $pinfo.Arguments = """$ShortcutPath"""
            $pinfo.WindowStyle = 'Hidden'
            $pinfo.CreateNoWindow = $true
            $p = New-Object -TypeName System.Diagnostics.Process
            $p.StartInfo = $pinfo
            $null = $p.Start()
            $p.WaitForExit()
  
            If($p.ExitCode -gt 0) 
            {
              Write-Warning -Message "Hashlinking failed for '$ShortcutPath'"
              Remove-Item $ShortcutPath -Force
              break
            } 
          }
          catch 
          {
            Write-Warning -Message 'Unable to create hashlnk'
            Write-Warning -Message "Error was $_"
            $line = $_.InvocationInfo.ScriptLineNumber
            Write-Warning -Message "Error was in Line $line"
            break
          }
          finally 
          {
            Stop-Process -Name Explorer
            $Results = 'Successful'
          }
        }
        else 
        {
          Write-Verbose -Message "Application path $ApplicationPath is invalid"
          break
        }
      }

      if($Remove.IsPresent) 
      {
        $ShortcutToRemove = Get-ChildItem -Path $GroupPath -Filter "$Name" -Recurse
        
        if(!([string]::IsNullOrEmpty($ShortcutToRemove))) 
        {
          if(Test-Path -Path $ShortcutToRemove.FullName) 
          {
            Write-Verbose -Message "Found $($ShortcutToRemove.FullName)"
          
            Try 
            {
              Remove-Item -Path $ShortcutToRemove.FullName -Force
            } 
            Catch 
            {
              Write-Warning -Message "Unable to remove shorcut like $Name"
              break
            }
            Finally 
            {
              Write-Verbose -Message "Removed $($ShortcutToRemove.FullName)"
              Stop-Process -Name explorer
              $Results = 'Successful'
            }
          }
        }
        else 
        {
          Write-Warning -Message "Unable to find shorcut like $Name"
          break
        }
      }
    
      Try 
      {
        If($LASTEXITCODE -gt 0) 
        {
          Write-Verbose "Last exit code was unhandled '$LASTEXITCODE'"
          throw
        }
      }
      Catch 
      {
        Write-Warning -Message "Error was $_"
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Warning -Message "Error was in Line $line"
        break
      }
      Finally 
      {
        if($Add.IsPresent) 
        {
          $Option = 'Add'
        }
        Else 
        {
          $Option = 'Remove'
        }
        
        if([string]::IsNullOrEmpty($Arguments)) 
        {
          $Arguments = 'None'
        }
        
        switch -Exact ($Option)
        {
          Add           
          { 
            $Result = New-Object -TypeName PSObject -Property @{
              Name            = $Name
              Option          = $Option
              ApplicationPath = $ApplicationPath
              Arguments       = $Arguments
              Group           = $Group
              Result          = $Results
              ElevateShortcut = $ElevateShortcut           
            }
          }
          Remove           
          {
            $Result = New-Object -TypeName PSObject -Property @{
              Name   = $Name
              Option = $Option
              Result = $Results
            }
          }
        }

        if(!([string]::IsNullOrEmpty($Result))) 
        {
          Write-Output -InputObject $Result
        }
        else 
        {
          Write-Verbose -Message 'No Results'
        }
      }
    }

    $Result = $null
    $Results = $null
    $VersionResult = $null
    $ScriptVersion = $null
    $GroupPath = $null
    $GroupPathName = $null
    $Name = $null
    $ApplicationPath = $null
    $ElevateShortcut = $null
    $Group = $null
    $Arguments = $null
    $Remove = $null
    $Add = $null
   
  }

  End 
  {
    Write-Verbose -Message "Ending $($MyInvocation.Mycommand)"
  }
}

#Set-WinXMenuItem -Add -ApplicationPath "$env:windir\system32\SnippingTool.exe" -Group Group3 -Name 'Snipping Tool' -ElevateShortcut

