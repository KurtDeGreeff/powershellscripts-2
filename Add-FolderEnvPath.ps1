Function Add-FolderEnvPath
{
  <#
    .SYNOPSIS
    Adds a path to the system enviroment path

    .DESCRIPTION
    Checks if the path is valid if so it will get the current enviroment system path and compares each path to the one which should be added if the path exists it will be added if not it will be skipped.

    .PARAMETER Version
    Output script version

    .PARAMETER PathToAdd
    Path to add

    .EXAMPLE
    Add-FolderEnvPath -Version
    Outputs the current script version to be used in future for automatic script updates

    .EXAMPLE
    Add-FolderEnvPath -PathToAdd C:\Tools\Utilities
    Adds the path to the system enviroment path

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
    [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline = $true, ParameterSetName = 'PathToAdd')]
    [string] $PathToAdd
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
      
      
        If(Test-Path -Path $PathToAdd) {
      
          Write-Verbose -Message "Path $PathToAdd is valid continuing"
      
        } else {
          Write-Warning -Message "Path $PathToAdd is not valid"
          break
        }

        Write-Verbose -Message "Adding folder $PathToAdd to SYSTEM path"
        
        [string[]]$PathArray = ($env:path).split(';')
        [bool]$PathAlreadyExists = $false
        [string]$FinalPath = $null
      
        Foreach ($Path in $PathArray) 
        {
          if($Path -ieq $PathToAdd) 
          {
            $PathAlreadyExists = $true
                
            Write-Verbose -Message "Path $Path is already added to system path"
          }
        }
      
        if(!($PathAlreadyExists))
        {
          $PathArray += $PathToAdd
          $FinalPath = ($PathArray -Join ';').replace(';;',';') # Fail safe in case there is a comma in the path passed to the script
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
        if(!([string]::IsNullOrEmpty($FinalPath))) 
        {
          $Result = New-Object -TypeName PSObject -Property @{
            Output = $FinalPath
          }
          [Environment]::SetEnvironmentVariable('path',$FinalPath, 'Machine')
        }
        else 
        {
          $Result = New-Object -TypeName PSObject -Property @{
            Output = 'No Changes Required'
          }
        }
        Write-Output -InputObject $Result
        
        $Result = $null
      }
         
      $VersionResult = $null
      $ScriptVersion = $null
      $NewPath = $null
      $PathArray = $null
    }

  }
  End
  {
    Write-Verbose -Message "Ending $($MyInvocation.Mycommand)"
  }
}
