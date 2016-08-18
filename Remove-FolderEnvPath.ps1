Function Remove-FolderEnvPath
{
  <#
      .SYNOPSIS
      Describe purpose of "Remove-FolderEnvPath" in 1-2 sentences.

      .DESCRIPTION
      Add a more complete description of what the function does.

      .PARAMETER Version
      Output script version

      .PARAMETER Path
      Path to remove

      .EXAMPLE
      Add-FolderEnvPath -Version
      Outputs the current script version to be used in future for automatic script updates

      .EXAMPLE
      Remove-FolderEnvPath -Path C:\Tools\Utilities\
      Removes the path to the system environment path

      .NOTES
      Author: Bevin Du Plessis

      .LINK
      https://github.com/nightshade2109/powershellscripts

      .INPUTS
      [string]

      .OUTPUTS
      [string]
  #>



  [Cmdletbinding()]
  param
  ( 
    [Parameter(ParameterSetName = 'Version')]
    [switch]$Version,
    [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline = $true, ParameterSetName = 'Path')]
    [String]$Path
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
      Try 
      {
        $EnvPaths = $null
        [string[]]$EnvPaths = ($env:path).split(';',[StringSplitOptions]::RemoveEmptyEntries)
        [bool]$PathExists = $false
        [string]$FinalEnvPath = $null
        [string[]]$PathsToCheck = $null
        [string]$MatchedPath = $null  
    
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
          
          
          Foreach ($PathToCheck in $PathsToCheck) 
          {
           Write-Verbose -Message "Comparing Path: $EnvPath to $PathToCheck"
            if($EnvPath -ieq $PathToCheck) 
            {
              $PathExists = $true
              $MatchedPath = $EnvPath
              Write-Verbose -Message "Path $MatchedPath is in the system environment path and will be removed"
            }
          }
        }
          
        if(!([string]::IsNullOrEmpty($MatchedPath)) -and $PathExists) 
        {
          $FinalEnvPath = ($EnvPaths -Join ';').Replace("$MatchedPath","$null").Replace(';;',';') # Fail safe in case there is a comma in the path passed to the script
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
    Write-Verbose -Message "Ending $($MyInvocation.Mycommand)"
  }
}
