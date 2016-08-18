Function Add-FolderEnvPath
{
  <#
      .SYNOPSIS
      Adds a path to the system environment path

      .DESCRIPTION
      Checks if the path is valid if so it will get the current environment system path and compares each path to the one which should be added if the path exists it will be added if not it will be skipped.

      .PARAMETER Version
      Output script version

      .PARAMETER Path
      Path to add

      .EXAMPLE
      Add-FolderEnvPath -Version
      Outputs the current script version to be used in future for automatic script updates

      .EXAMPLE
      Add-FolderEnvPath -Path C:\Tools\Utilities
      Adds the path to the system environment path

      .NOTES
      Author: Bevin Du Plessis

      .LINK
      https://github.com/nightshade2109/powershellscripts

      .INPUTS
      [string]

      .OUTPUTS
      [string]
  #>


  [CmdletBinding()]
  Param
  (    
    [Parameter(ParameterSetName = 'Version')]
    [switch]$Version,
    [Parameter(Mandatory = $true,HelpMessage = 'Path to add',Position = 0,ValueFromPipeline = $true, ParameterSetName = 'Path')]
    [string] $Path
  )

  Begin
  {
    Write-Verbose -Message "Starting $($MyInvocation.Mycommand)"
    
    $ScriptVersion = '0.1'

    if ($Version.IsPresent) 
    {
      $VersionResult = New-Object -TypeName PSObject -Property @{
        Version = $ScriptVersion
      }
      
      Write-Output -InputObject $VersionResult
    }
    $VersionResult = $null
    $ScriptVersion = $null

    if ($PSBoundParameters.ContainsKey('Debug'))
    {
      $DebugPreference = 'Continue'
    }

    if ($PSBoundParameters.ContainsKey('Verbose'))
    {
      $VerbosePreference = 'Continue'
    }
  }
  
  Process
  {
    if(!($Version.IsPresent)) 
    {
      Write-Verbose -Message "Processing $($MyInvocation.Mycommand)"
      Try
      {
        If(Test-Path -Path $Path) 
        {
          Write-Verbose -Message "Path $Path is valid continuing"
        } 
        else 
        {
          Write-Warning -Message "Path $Path is not valid"
          break
        }

        Write-Verbose -Message "Adding folder $Path to system environment path"
        
        $EnvPaths = $null
        
        [string[]]$EnvPaths = ($env:path).split(';')
        [bool]$PathAlreadyExists = $false
        [string]$FinalEnvPath = $null
        
            
        # Build possible combinations of paths to check with or without slashes
        $PathLastCharacter = $Path[-1]
        if($PathLastCharacter -ieq [IO.Path]::DirectorySeparatorChar -or $PathLastCharacter -ieq [IO.Path]::AltDirectorySeparatorChar) 
        {
          $PathsToCheck += $Path.Substring(0,$Path.Length-1)
          $PathsToCheck += $Path
        }
        else 
        {
          $PathsToCheck += $(Join-Path -Path $Path -ChildPath '\')
          $PathsToCheck += $Path
        }
        
        
        
      
        Foreach ($EnvPath in $EnvPaths) 
        {
        
          Foreach ($PathToCheck in $PathsToCheck) {
          
            if($EnvPath -ieq $PathToCheck) 
            {
              $PathAlreadyExists = $true
              Write-Verbose -Message "Path $EnvPath is already added to system path"
              break
            }
          
          }
   
        }
      
        if(!($PathAlreadyExists))
        {
          $EnvPaths = $null
          [char]$Char = ';'
        
        
          $EnvPathLastCharacter = $env:path[-1]
          if($EnvPathLastCharacter -eq $Char) 
          {
            Write-Verbose -Message 'Removing ; character from path'
            $NewEnvPaths = $env:path.Substring(0,$env:path.Length-1)
            Write-Verbose -Message "Updated path: $NewEnvPaths"
          }
        
          [string[]]$EnvPaths = ($NewEnvPaths).split(';')
                
          $EnvPaths += $Path
          $FinalEnvPath = ($EnvPaths -Join ';').replace(';;',';') # Fail safe in case there is a comma in the path passed to the script
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
        $Result = $null
        
        if(!([string]::IsNullOrEmpty($FinalEnvPath))) 
        {
          $Result = New-Object -TypeName PSObject -Property @{
            Output = $FinalEnvPath
          }
          [Environment]::SetEnvironmentVariable('Path',$FinalEnvPath, 'Machine')
        }
        else 
        {
          $Result = New-Object -TypeName PSObject -Property @{
            Output = 'No Changes Required'
          }
        }
        Write-Output -InputObject $Result
      }
    }

  }
  End
  {
    Update-EnvVarsForSession
    Write-Verbose -Message "Ending $($MyInvocation.Mycommand)"
  }
}
