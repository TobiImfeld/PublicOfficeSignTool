#
# Script.ps1
#
#Win SignTool help-link: https://learn.microsoft.com/en-us/dotnet/framework/tools/signtool-exe

[System.IO.DirectoryInfo]$Script:path_signtool = [System.IO.Path]::GetFullPath( (join-Path -Path $PSScriptRoot "\signtool.exe") )

function Add-Signature {
    param(
        [String] $Filename,
        [String] $Certificate,
        [string] $Passphrase
    )

    # Check is signtool in script root path
    if (-not [System.IO.DirectoryInfo]$Script:path_signtool){
        
        # Use environmental setting for Sigtool if set
        if ( (Test-Path 'env:SIGNTOOL') ){
            [System.IO.DirectoryInfo]$Script:path_signtool = $env:SIGNTOOL + "signtool.exe"
        }

        if (-not (Test-Path $Script:path_signtool) ){
            Write-Error "SignTool is missing!"
            return $false
        }
    }

    # Get the path to the Office file to work on
    [System.IO.DirectoryInfo]$Script:path_officefile  = [System.IO.Path]::GetFullPath( $Filename )

    [System.IO.DirectoryInfo]$Script:path_certificate = [System.IO.Path]::GetFullPath( $Certificate )

    # Get the path to the Office file to work on
    if (-not (Test-Path $Script:path_officefile) ){
        Write-Warning "The office file [ $Script:path_officefile ] does not exist!"
        return $false
    }

    if (-not (Test-Path $Script:path_certificate) ){
        Write-Warning "The certificate [ $Script:path_certificate ] cannot be found!"
        return $false
    }

    if (-not (Test-Path $Script:path_signtool)){
        Write-Warning "The [ $Script:path_signtool ] cannot be found!"
        return $false
    }

    if ((Test-Path $Script:path_signtool) -and 
        (Test-Path $Script:path_officefile) -and 
        (Test-Path $Script:path_certificate) 
    ){
        #$command = "Sign /q /f '$Script:path_certificate' /p `"$Passphrase`" /fd `"SHA256`" /td `"SHA256`" `"$Script:path_officefile`" ";
        $command = "`"$Script:path_signtool`" Sign /q /f `"$Script:path_certificate`" /p `"$Passphrase`" /fd `"SHA256`" /td `"SHA256`" `"$Script:path_officefile`" ";
        #$command = '"{0}" Sign /q /f "{1}" /p "{2}" /fd "{3}" /td "{3}" "{4}" ' -f $Script:path_signtool, $Script:path_certificate, $Passphrase, 'SHA256', $Script:path_officefile;
        $success = $null
        $cmdOutput = cmd /c $command '2>&1' | ForEach-Object { #Create a handle for the signtool.exe?!
            if ($_ -is [System.Management.Automation.ErrorRecord]) {
                $success = $false
                Write-Error $_
            } else {
                if ($_ -like '*PFX password is not correct.'){
                    Write-Warning "Incorrect Password [ $Passphrase ]"
                    $success = $false
                } elseif ($_ -like 'SignTool Error:*'){
                    Write-Warning "Unable to Sign [ $Script:path_officefile ]"
                    $success = $false
                } elseif ($_ -like 'Error information:*'){
                    Write-Warning "It is likely that this file contains no Macros."
                    $success = $false
                } elseif ($_ -like '*Adding Additional Store'){
                    $success = $true
                    Write-Host "Sucessfully signed file [ $Script:path_officefile ]"
#                    #return $true
                } else {
                    write-Host $_
                }
            }
        }
        return $success
    } else {
        write-error("General Failure!")
    }
    return $false
}


function Get-Signature {
    param(
        [String] $Filename
    )

    # Use environmental setting for Sigtool if set
    if ( (Test-Path env:SIGNTOOL) ){
        [System.IO.DirectoryInfo]$Script:path_signtool = $env:SIGNTOOL + "signtool.exe"
    }

    # Get the path to the Office file to work on
    [System.IO.DirectoryInfo]$Script:path_officefile  = [System.IO.Path]::GetFullPath( (Join-Path -path $pwd $Filename) )

    if (-not (Test-Path $Script:path_officefile) ){
        Write-Warning "The office file [ $Script:path_officefile ] does not exist!"
        return $false
    }

    if ((Test-Path $Script:path_signtool) -and 
        (Test-Path $Script:path_officefile)
    ){
        $command = "verify `"$Script:path_officefile`" ";
        $cmdOutput = cmd /c "`"$Script:path_signtool`" $command" '2>&1' | ForEach-Object {
            write-host $_
        }
    } else {
        write-error("General Failure!")
    }
}

#Export-ModuleMember -Function Add-Signature
#Export-ModuleMember -Function Get-Signature
