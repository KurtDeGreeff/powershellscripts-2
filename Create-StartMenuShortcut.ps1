Function Create-StartMenuShortcut
{
  <#
    .SYNOPSIS
    Create a start menu shortcut

    .DESCRIPTION
    Create a start menu shortcut using the WScript.Shell comobject

    .PARAMETER Version
    Output version of this script

    .PARAMETER Target
    This property is for the shortcut's target path only. Any arguments to the shortcut must be placed in the Argument's property.
    Path to the target application for example C:\Tools\Utilities\Cmder\Cmder.exe

    .PARAMETER AllUsers
    Create the shortcut for all users

    .PARAMETER Folder
    Folder to in the start menu to create the shortcut if the folder does not exist it will be created

    .PARAMETER Name
    Name of the shortcut

    .PARAMETER WorkingDirectory
    Assign a working directory to a shortcut, or identifies the working directory used by a shortcut.

    .PARAMETER Arguments
    The arguments parameter can contain specific arguments to pass to the application being started by the shortcut.
    For example you can make the target internet explorer 'C:\Program Files\Internet Explorer\iexplore.exe' then pass a URL with the argument parameter 'www.microsoft.com'
   

    .PARAMETER Icon
    A string that locates the icon. The string should contain a fully qualified path
    Icon location this can either be an executable such as an exe, dll or a ico file.

    .PARAMETER IconIndex
    this is useful if the icon file has multiple icons you can specific the index of the icon to use this defaults to zero so it will use the first icon

    .PARAMETER Description
    The Description parameter contains a string value describing the shortcut.

    .PARAMETER WindowStyle
    Sets the window style for the program being run
    1 Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position.
    3 Activates the window and displays it as a maximized window.
    7 Minimizes the window and activates the next top-level window.

    .PARAMETER HotKey
    Assigns a key-combination to a shortcut, or identifies the key-combination assigned to a shortcut.
    For example CTRL+SHIFT+F1 explorer will need to be restarted in order for the hotkey to take effect.

    .PARAMETER RelativePath
    Assigns a relative path to a shortcut, or identifies the relative path of a shortcut.

    .PARAMETER ElevateShortcut
    Set the shortcut to always run as administrator

    .PARAMETER Force
    Overwrite existing shortcut

    .EXAMPLE
    Create-StartMenuShortcut -Version
    Describe what this call does

    .EXAMPLE
    Create-StartMenuShortcut -Name "Cmder" -Target "C:\Tools\Utilities\Cmder\Cmder.exe" -Folder "Cmder" -WorkingDirectory "C:\Tools" -Icon "C:\Tools\Utilities\Cmder\Cmder.exe" -IconIndex 0 -Description "Cmder Alternative Shell" -HotKey "CTRL+SHIFT+F1" -WindowStyle 1 -RelativePath "C:\Tools" -Force -Verbose -Debug

    .NOTES
    Author: Bevin Du Plessis

    More info :https://msdn.microsoft.com/en-us/library/xk6kst2k(v=vs.84).aspx

    .LINK
    https://github.com/nightshade2109/powershellscripts
#>


  [CmdletBinding()]
  param(
    [Parameter(ParameterSetName = 'Version')]
    [switch]$Version,
    [Parameter(Mandatory = $true)][string]$Target,
    [switch]$AllUsers,
    [Parameter(Mandatory = $true)][string]$Folder,
    [Parameter(Mandatory = $true)][string]$Name,
    [string]$WorkingDirectory = $null,
    [string]$Arguments = $null,
    [string]$Icon = $null,
    [int]$IconIndex = 0,
    [string]$Description = $null,
    
    # 1 Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position.
    # 3 Activates the window and displays it as a maximized window.
    # 7 Minimizes the window and activates the next top-level window.
    [int]$WindowStyle = 1,
    
    # Can be any one of the following: ALT+, CTRL+, SHIFT+
    # a ... z, 0 ... 9, F1 F12, ...
    # The KeyName is not case-sensitive.
    [string]$HotKey = $null,
    
    [string]$RelativePath = $null,
    [bool]$ElevateShortcut = $false,
    [switch]$Force
  )

  Begin {
    Set-StrictMode -Version Latest
    
    # Pass verbose and debug to child functions
    if ($PSBoundParameters.ContainsKey('Debug'))
    {
      $script:DebugPreference = 'Continue'
    }
    else 
    {
      $script:VerbosePreference = 'SilentlyContinue'
    }

    if ($PSBoundParameters.ContainsKey('Verbose'))
    {
      $script:VerbosePreference = 'Continue'
    }
    else 
    {
      $script:VerbosePreference = 'SilentlyContinue'
    }
  
    #############################################################
    Write-Verbose -Message "Starting $($MyInvocation.Mycommand)"
    #############################################################

  }

  Process
  {

    $ScriptVersion = '0.1'

    if ($Version.IsPresent) 
    {
      $VersionResult = New-Object -TypeName PSObject -Property @{
        Version = $ScriptVersion
      }
      
      Write-Output -InputObject $VersionResult
    }

        
        
    if (!$Version.IsPresent) 
    {
      #############################################################
      Write-Verbose -Message "Processing $($MyInvocation.Mycommand)"
      #############################################################
    
      $Result = $null
      $Results = $null
      
      
      Try
      {
        If($AllUsers) 
        {
          $BasePath = "$env:AllUsersProfile\Microsoft\Windows"
        } 
        Else 
        {
          $BasePath = "$env:USERPROFILE"
        }
        
        $Shortcut = "$BasePath\Start Menu\Programs\$Folder\$Name.lnk"
        
        if((Test-Path -Path "$Shortcut") -and !($Force.IsPresent)) 
        {
          Write-Verbose -Message "Shortcut $Shortcut exists"
          $Results = 2
        }
        else 
        {
          if($Force.IsPresent) 
          {
            Remove-Item -Path "$BasePath\Start Menu\Programs\$Folder" -Recurse -Force -Confirm:$false
          }
        
          if(!(Test-Path -Path "$BasePath\Start Menu\Programs\$Folder")) 
          {
            Write-Verbose "Creating Folder '$BasePath\Start Menu\Programs\$Folder'"
            New-Item -Path "$BasePath\Start Menu\Programs\$Folder" -ItemType Directory -Force
          }
        
          $WshShell = New-Object -ComObject WScript.Shell
          $objShortCut = $WshShell.CreateShortcut($Shortcut)
          $objShortCut.TargetPath = $Target
          
          if($Arguments -ne $null) 
          {
            $objShortCut.Arguments = "$Arguments"
          }
          
          if($Icon -ne $null) 
          {
            if($IconIndex -ne 0) 
            {
              $objShortCut.IconLocation = "$Icon,$IconIndex"
            }
            else 
            {
              $objShortCut.IconLocation = "$Icon,0"
            }
          }
          
          if($Description -ne $null) 
          {
            $objShortCut.Description = "$Description"
          }
          
          if($WorkingDirectory -ne $null) 
          {
            $objShortCut.WorkingDirectory = "$WorkingDirectory"
          }
         
          if($WindowStyle -ne 1) 
          {
            $objShortCut.WindowStyle = $WindowStyle
          }
               
          if($HotKey -ne $null) 
          {
            Write-Verbose -Message 'Shorcut hotkey provided explorer will need to be restarted for it to take effect!'
            $objShortCut.Hotkey = "$HotKey"
          }
          
          if($RelativePath  -ne $null) 
          {
            $objShortCut.RelativePath  = "$RelativePath"
          }
           
          $objShortCut.Save()
          
          
          
          if($ElevateShortcut) 
          {
            $bytes = [IO.File]::ReadAllBytes("$Shortcut")
            $bytes[0x15] = $bytes[0x15] -bor 0x20 #set byte 21 (0x15) bit 6 (0x20) ON
            [IO.File]::WriteAllBytes("$Shortcut", $bytes)
          }
          
          if(Test-Path -Path "$Shortcut") 
          {
            $Results = 0
            Write-Verbose -Message "Shortcut $Shortcut created successfully"
          } 
          Else 
          {
            Write-Verbose -Message "Creating shortcut $Shortcut was unsuccessful"
            $Results = 1
          }
        }
      }
      Catch 
      {
        Write-Warning -Message "Error was $_"
        $line = $_.InvocationInfo.ScriptLineNumber
        Write-Warning -Message "Error was in Line $line"
        
        if(Test-Path -Path "$BasePath\Start Menu\Programs\$Folder") 
        {
          Write-Verbose -Message "Removing Folder '$BasePath\Start Menu\Programs\$Folder'"
          Remove-Item -Path "$BasePath\Start Menu\Programs\$Folder" -Force
        }
        
        break
      }
      Finally 
      {
        if(!([string]::IsNullOrEmpty($Results))) 
        {
          If($Results -eq 0) 
          {
            $Result = New-Object -TypeName PSObject -Property @{
              Output = 'Successful'
            }
          }
          elseif ($Results -eq 1) 
          {
            $Result = New-Object -TypeName PSObject -Property @{
              Output = 'Failed'
            }
          }
          elseif ($Results -eq 2) 
          {
            $Result = New-Object -TypeName PSObject -Property @{
              Output = 'Shortcut already exists use -Force parameter to overwrite'
            }
          }
          else 
          {
            $Result = New-Object -TypeName PSObject -Property @{
              Output = 'Unhandled Error Code Returned'
            }
          }
        }
        else 
        {
          $Result = New-Object -TypeName PSObject -Property @{
            Output = 'No Results Returned'
          }
        }
        Write-Output -InputObject $Result
      }
    }

    #############################################################
  }
  

  End
  {
    if (!$Version.IsPresent) 
    {
      #############################################################
      Write-Verbose -Message "Ending $($MyInvocation.Mycommand)"
      #############################################################
    }
  }
}
#Create-StartMenuShortcut -Name "Cmder" -Target "C:\Tools\Utilities\Cmder\Cmder.exe" -Folder "Cmder" -WorkingDirectory "C:\Tools" -Icon "C:\Tools\Utilities\Cmder\Cmder.exe" -IconIndex 0 -Description "Cmder Alternative Shell" -HotKey "CTRL+SHIFT+F1" -WindowStyle 1 -RelativePath "C:\Tools" -Force -Verbose -Debug
