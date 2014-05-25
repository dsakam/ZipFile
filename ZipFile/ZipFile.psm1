######## LICENSE ####################################################################################################################################
<#
 # Copyright (c) 2013-2014, Daiki Sakamoto
 # All rights reserved.
 #
 # Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
 #   - Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
 #   - Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer
 #     in the documentation and/or other materials provided with the distribution.
 #
 # THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
 # THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
 # HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
 # LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
 # ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE
 # USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 #
 #>
 # http://opensource.org/licenses/BSD-2-Clause
#####################################################################################################################################################

######## HISTORY ####################################################################################################################################
<#
 # Zip File Compression / Decompression Module for PowerShell
 #
 #  2013/11/15  Version 0.0.0.1
 #  2013/12/12  Version 1.0.0.0  1st Public Release
 #  2013/12/13  Version 1.0.0.1  File Delete Message is changed from Waning into Verbose.
 #  2013/12/17  Version 1.0.0.2  Minor Change
 #  2013/12/28  Version 1.0.0.3  No change is made to this script.
 #  2014/01/06  Version 1.0.1.0  Add Data Check of "InputObject" Parameter of Expand-Zip Comdlet
 #  2014/01/06  Version 1.0.2.0  Change type of a exception from "System.IO.InvalidDataException" to "System.IO.FileFormatException"
 #                               when content of file ("InputObject" Parameter) is empty.
 #  2014/01/10  Version 1.0.3.0  Add validation of "-InputObject" paramater of Expand-ZipFile Cmdlet.
 #  2014/01/14  Version 1.0.4.0  Change type of Zip Header from [object] into [byte[]] of New-ZipFile Cmdlet.
 #  2014/01/17  Version 1.0.5.0  Change expression of Type check
 #  2014/02/23  Version 1.1.0.0  Change some double quotations (") into single quotations (').
 #                               Reconsidered style of verbose output.
 #                               (New-ZipFile) Return [string]::Empty, when Zip compression is aborted because the file already exists.
 #  2014/02/27  Version 1.1.1.0  Modify verbose output style.
 #                               (New-ZipFile) Change from Write-Verbose into Write-Warning when the file already exists.
 #  2014/02/27  Version 1.1.2.0  Change convert some double quotations (") and single quotations (').
 #  2014/05/15  Version 1.2.0.0  Change validation of parameter of 'Expand-ZipFile' Cmdlet because of huge memory exhaust.
 #  2014/05/22  Version 1.3.0.0  Fix bug regarding $Path parameter of 'New-ZipFile' cmdlet.
 #                               Change indent style.
 #                               Major change of internal process, especially regarding exception.
 #  2014/05/24  Version 1.3.1.0  Add Help content about 'Path' parameter.
 #                               Fix bug regarding output path of 'New-ZipFile' cmdlet.
 #
 #>
#####################################################################################################################################################

#####################################################################################################################################################
# Variables
$script:AssemblyName = 'System.IO.Compression.FileSystem, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
<#
$script:AssemblyName = 'System.IO.Compression, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
#>

#####################################################################################################################################################
Function Expand-ZipFile
{
    <#
        .SYNOPSIS
            Zip ファイルを解凍します。

        .DESCRIPTION
            Zip ファイルのファイル名から拡張子 (.zip) を除いたフォルダを作成し、その中に Zip ファイルを解凍します。

            Expand-ZipFile コマンドレットは、まず System.IO.Compression.FileSystem.dll のロードを試みます。
            System.IO.Compression.FileSystem.dll のロードに成功すると、System.IO.Compression.ZipFile クラスの 
            ExtractToDirectory メソッドを使って Zip ファイルを解凍します。
    
            Force パラメーターが指定されていないときは、System.IO.Compression.ZipArchive クラスを使って 
            Zip ファイル内のエントリーをチェックします。
            これは Zip ファイルに含まれるファイルの数が多い場合はパフォーマンスに影響する可能性があるので、
            そのような場合は Force オプションを指定することを検討してください。

            System.IO.Compression.FileSystem.dll のロードに失敗すると、シェルモードで Zip ファイルを解凍します。
            シェルモードでは、System.Shell.Folder.CopyHere メソッドを使って Zip ファイルを解凍します。

        .PARAMETER InputObject
            解凍する Zip ファイルのパスを指定します。

        .PARAMETER Path
            解凍先のフォルダーのパスを指定します。

            このパラメーターで指定されたフォルダーに、InputObject パラメーターから拡張子 (.zip) を除いた名前のフォルダーが
            自動的に作成され、そのフォルダーの中に Zip ファイルが解凍されます。

        .PARAMETER ShellMode
            強制的にシェルモードで実行するときに指定します。

        .PARAMETER Force
            解凍先のフォルダーが既に存在する場合、そのフォルダを削除してから Zip ファイルを解凍します。

        .PARAMETER Verbose
            詳細情報を表示します。

        .INPUTS
            System.String
            パイプを使用して、InputObject パラメーターを渡すことができます。

        .OUTPUTS
            System.String
            Zip ファイルを解凍したフォルダのパスを返します。
            Zip ファイルが解凍されなかった場合は System.String.Empty を返します。

        .EXAMPLE
            Expand-ZipFile sample.zip
            カレントディレクトリにある sample.zip ファイルをカレントディレクトリの sample フォルダに解凍します。

        .EXAMPLE
            Expand-ZipFile -InputObject .\sample.zip -Path C:\Temp
            カレントディレクトリにある sample.zip ファイルを C:\Temp\sample フォルダに解凍します。

        .NOTES
            System.IO.Compression.ZipFile は .NET Framework 4.5 からサポートされています。

        .LINK
            copyHere Method (System.Shell.Folder)
            http://msdn.microsoft.com/ja-jp/library/windows/desktop/ms723207.aspx

            Compress Files with Windows PowerShell then package a Windows Vista Sidebar Gadget - David Aiken - Site Home - MSDN Blogs
            http://blogs.msdn.com/b/daiken/archive/2007/02/12/compress-files-with-windows-powershell-then-package-a-windows-vista-sidebar-gadget.aspx

            CopyHere メソッドから Zip ファイルを処理することはできません
            http://support.microsoft.com/kb/2679832

            Assembly.Load メソッド (AssemblyName) (System.Reflection)
            http://msdn.microsoft.com/ja-jp/library/x4cw969y.aspx

            ZipFile クラス (System.IO.Compression)
            http://msdn.microsoft.com/ja-jp/library/system.io.compression.zipfile.aspx

            ZipArchive クラス (System.IO.Compression)
            http://msdn.microsoft.com/ja-jp/library/system.io.compression.ziparchive.aspx
    #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [ValidateScript ( {
            # File Existence Check
            if (-not (Test-Path -Path $_)) { throw New-Object System.IO.FileNotFoundException }

            # Check File or Directory / [+]V1.0.3.0 (2014/01/10) / [*]V1.0.5.0 (2014/01/17)
            if ((Get-Item -Path $_) -isnot [System.IO.FileInfo]) { throw New-Object System.IO.FileNotFoundException }

            # Data Check [*]V1.2.0.0 (2014/05/15)
            if (($filehead = Get-Content -Path $_ -Encoding Byte -First 2) -eq $null) { throw New-Object System.IO.FileFormatException }

            # File Format Check [*]V1.2.0.0 (2014/05/15)
            if ([System.Text.Encoding]::ASCII.GetString($filehead) -ne 'PK')
            {
                throw New-Object System.IO.FileFormatException
            }

            return $true
        } )]
        [string]$InputObject,

        [Parameter(Mandatory=$false, Position=1)]
        [ValidateScript ( {
            if (-not (Test-Path -Path $_)) { New-Item -Path $_ -ItemType Directory }
            return $true
        } )]
        [string]$Path = ($InputObject | Resolve-Path | Split-Path -Parent),

        [Parameter(Mandatory=$false)][switch]$ShellMode,
        [Parameter(Mandatory=$false)][switch]$Force
    )

    Process
    {
        # Remove the following flag / [-]V1.3.0.0 (2014/05/23)
        # [bool]$completed = $false

        # Add the following flag / [+]V1.3.0.0 (2014/05/23)
        [bool]$assembly_loaded = $false

        $source_Path = ($InputObject | Convert-Path)
        $source_Filename = ($source_Path | Split-Path -Leaf)
        $destination_Path = ($Path | Resolve-Path | Join-Path -ChildPath ([System.IO.FileInfo]$InputObject).BaseName)


        if (-not $ShellMode)
        {
            try
            {
                # Load Assembly / Add 'if' condition [*]V1.3.0.0 (2014/05/23)
                if ([System.Reflection.Assembly]::Load($script:AssemblyName) -ne $null)
                {
                    Write-Verbose ('[' + $MyInvocation.MyCommand.Name + ']' + " '" + ($script:AssemblyName -split ',')[0] + "' is loaded successfully." )

                    # [+]V1.3.0.0 (2014/05/23)
                    $assembly_loaded = $true
                }


                if ($Force)
                {
                    # Force
                    if ($destination_Path | Test-Path) { $destination_Path | Get-ChildItem -Force -Recurse | Remove-Item -Force -Recurse }
                }
                else
                {
                    # Check Zip Entries
                    [System.IO.Compression.ZipArchive]$archive = [System.IO.Compression.ZipFile]::OpenRead($source_Path)

                    $archive.Entries | % {

                        Write-Verbose ('[' + $MyInvocation.MyCommand.Name + ']' + " ENTRY: $source_Filename/" + ([System.IO.Compression.ZipArchiveEntry]$_).FullName)

                        if (($entry = ($destination_Path | Join-Path -ChildPath ([System.IO.Compression.ZipArchiveEntry]$_).FullName)) | Test-Path)
                        {
                            # [*]V1.0.5.0 (2014/01/17)
                            if ((Get-Item -Path $entry) -isnot [System.IO.DirectoryInfo])
                            {
                                # File Already Exist (Remove entry...)
                                Remove-Item -Path $entry -Force
                                Write-Verbose ('[' + $MyInvocation.MyCommand.Name + ']' + " REMOVE: $entry")
                            }
                        }
                    }
                }
            
                # Unzip (by .NET)
                [System.IO.Compression.ZipFile]::ExtractToDirectory($source_Path, $destination_Path)
                
                Write-Verbose ('[' + $MyInvocation.MyCommand.Name + ']' + " '$source_Filename' -> '$destination_Path'")
                
                # [-]V1.3.0.0 (2014/05/23)
                # $completed = $true
            }
            catch  # [+]V1.3.0.0 (2014/02/23)
            {
                if (-not $assembly_loaded)
                {
                    Write-Warning ('[' + $MyInvocation.MyCommand.Name + ']' + " Fail to load '" + ($script:AssemblyName -split ',')[0] + "'!" )
                }

                $destination_Path = [string]::Empty
            }
            finally
            {
                if ($archive -ne $null) { $archive.Dispose() }
            }
        }


        # Change condition / [*]V1.3.0.0 (2014/05/23)
        if (-not $assembly_loaded)
        {
            # Shell-Mode
            Write-Verbose ('[' + $MyInvocation.MyCommand.Name + '](Shell)' + " Enter Shell mode...")

            # Destination Folder
            if (-not ($destination_Path | Test-Path)) { New-Item -Path $destination_Path -ItemType Directory }

            $shell = New-Object -ComObject Shell.Application
            $zip = $shell.NameSpace($source_Path)
            $dest = $shell.NameSpace($destination_Path)
            $opt = 0

            # Verbose
            if ($VerbosePreference -eq 'SilentlyContinue') { $opt += (4 + 1024) }

            # Force
            if ($Force) { $opt += 16 }

            # Unzip (by Shell)
            $dest.CopyHere($zip.Items(), $opt)
            Write-Verbose ('[' + $MyInvocation.MyCommand.Name + '](Shell)' + " '$source_Filename' -> '$destination_Path'")
        }


        # Output -> Return / [*]V1.3.0.0 (2014/05/23)
        return $destination_Path
    }
}

#####################################################################################################################################################
Function New-ZipFile
{
    <#
        .SYNOPSIS
            Zip ファイルを作成します。

        .DESCRIPTION
            指定されたフォルダーまたはファイルから、フォルダー名またはファイル名に拡張子 .zip を付加した Zip ファイルを作成します。

            New-ZipFile コマンドレットは、まず System.IO.Compression.FileSystem.dll のロードを試みます。
            System.IO.Compression.FileSystem.dll のロードに成功すると、圧縮対象がフォルダーの場合は 
            System.IO.Compression.ZipFile クラスの CreateFromDirectory メソッドを使って Zip ファイルを作成します。
            圧縮対象がファイルの場合は、System.IO.Compression.ZipArchive クラス、および System.IO.Compression.ZipFile クラスと 
            System.IO.Compression.ZipFileExtensions クラスの CreateEntryFromFile 拡張メソッドを使って Zip ファイルを作成します。

            System.IO.Compression.FileSystem.dll のロードに失敗すると、シェルモードで Zip ファイルを作成します。
            シェルモードでは、System.Shell.Folder.CopyHere メソッドを使って Zip ファイルを作成します。
            シエルモードでは、フォルダーの圧縮をサポートしていません。

        .PARAMETER InputObject
            Zip 圧縮するフォルダーまたはファイルのパスを指定します。

        .PARAMETER Path
            作成した Zip ファイルを保存するフォルダーのパスを指定します。

            InputObject パラメーターがファイルの場合、作成される Zip ファイルには InputObject の拡張子を '.zip' に変更したファイル名が
            自動的につけられます。
            InputObject パラメーターがフォルダーの場合は InputObject に拡張子 '.zip' を付加したファイル名になります。

            いずれの場合も、Zip ファイルのパスが InputObject と同じになる場合は、エラーになります。

            省略した場合は、InputObject が置かれているフォルダーに Zip ファイルを作成します。

        .PARAMETER ShellMode
            強制的にシェルモードで実行するときに指定します。
            ただし、シェルモードはフォルダーの圧縮をサポートしていません。
            そのため、シェルモードで InputObject がフォルダーの場合、New-ZipFile コマンドレットは System.NotSupportedException 
            をスローして終了します。

        .PARAMETER Force
            作成する Zip ファイルのパスに既にファイルが存在する場合、そのファイルを削除してから Zip ファイルを作成します。
            このパラメーターを指定しないと、作成する Zip ファイルのパスに既にファイルが存在する場合、New-ZipFile コマンドレットは
            何もせずに終了します。

        .PARAMETER Verbose
            詳細情報を表示します。

        .INPUTS
            System.String
            パイプを使用して、InputObject パラメーターを渡すことができます。

        .OUTPUTS
            System.String
            作成した Zip ファイルのパスを返します。
            Zip ファイルが作成されなかった場合は System.String.Empty を返します。

        .EXAMPLE
            New-ZipFile .\test.txt
            カレントディレクトリにある test.txt を Zip 圧縮して test.zip を作成します。

        .EXAMPLE
            New-ZipFile -InputObject .\test.txt -Path C:\Temp
            カレントディレクトリにある test.txt ファイルを Zip 圧縮して C:\Temp\test.zip ファイルを作成します。

        .NOTES
            System.IO.Compression.ZipFile は .NET Framework 4.5 からサポートされています。

        .LINK
            copyHere Method (System.Shell.Folder)
            http://msdn.microsoft.com/ja-jp/library/windows/desktop/ms723207.aspx

            Compress Files with Windows PowerShell then package a Windows Vista Sidebar Gadget - David Aiken - Site Home - MSDN Blogs
            http://blogs.msdn.com/b/daiken/archive/2007/02/12/compress-files-with-windows-powershell-then-package-a-windows-vista-sidebar-gadget.aspx

            CopyHere メソッドから Zip ファイルを処理することはできません
            http://support.microsoft.com/kb/2679832

            Assembly.Load メソッド (AssemblyName) (System.Reflection)
            http://msdn.microsoft.com/ja-jp/library/x4cw969y.aspx

            ZipFile クラス (System.IO.Compression)
            http://msdn.microsoft.com/ja-jp/library/system.io.compression.zipfile.aspx

            ZipArchive クラス (System.IO.Compression)
            http://msdn.microsoft.com/ja-jp/library/system.io.compression.ziparchive.aspx
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [ValidateScript ( {
            if (-not (Test-Path -Path $_)) { throw New-Object System.IO.FileNotFoundException }
            return $true
        } )]
        [string]$InputObject,

        [Parameter(Mandatory=$false, Position=1)]
        [ValidateScript ( {
            if (-not (Test-Path -Path $_)) { New-Item -Path $_ -ItemType Directory }

            # [+]V1.3.1.0 (2014/05/25)
            if ((Get-Item -Path $_) -isnot [System.IO.DirectoryInfo]) { throw New-Object System.IO.DirectoryNotFoundException }

            return $true
        } )]
        [string]$Path = ($InputObject | Resolve-Path | Split-Path -Parent),

        [Parameter(Mandatory=$false)][switch]$ShellMode,
        [Parameter(Mandatory=$false)][switch]$Force
    )

    Process
    {
        # Remove the following flags / [-]V1.3.0.0 (2014/05/23)
        # [bool]$aborted = $false
        # [bool]$completed = $false

        # Add the following flag / [+]V1.3.0.0 (2014/05/23)
        [bool]$assembly_loaded = $false

        $source_Path = ($InputObject | Convert-Path)
        $source_Filename = ($source_Path | Split-Path -Leaf)


        # Zip File Name (Path) / [*]V1.0.5.0 (2014/01/17)
        if ((Get-Item -Path $source_Path) -is [System.IO.FileInfo])
        {
            # File
            $destination_Path = ($Path | Resolve-Path | Join-Path -ChildPath ([System.IO.FileInfo]$source_Path).BaseName) + '.zip'
        }
        else
        {
            # Directory
            $destination_Path = ($Path | Resolve-Path | Join-Path -ChildPath ($source_Path | Split-Path -Leaf)) + '.zip'
        }


        # Validation of Destination Path / [+]V1.3.0.0 (2014/05/23)
        if ($source_Path -eq $destination_Path)
        {
            throw New-Object System.ArgumentException "Destination Path ('$source_Path') is same as the original zip file."
        }

        # Validation of Destination Path / [+]V1.3.1.0 (2014/05/25)
        if ((Get-Item -Path $source_Path) -is [System.IO.DirectoryInfo])
        {
            if ($destination_Path.Length -gt $source_Path.Length)
            {
                if ((Split-Path -Path $destination_Path -Parent) -contains $source_Path)
                {
                    throw New-Object System.ArgumentException "Destination Path ('$source_Path') is same as the original zip file."
                }
            }
        }
        #>


        # Decompression by .NET Framework
        if (-not $ShellMode)
        {
            try
            {
                # Load Assembly / Add 'if' condition [*]V1.3.0.0 (2014/05/23)
                if ([System.Reflection.Assembly]::Load($script:AssemblyName) -ne $null)
                {
                    Write-Verbose ('[' + $MyInvocation.MyCommand.Name + ']' `
                        + " '" + ($script:AssemblyName -split ',')[0] + "' is loaded successfully." )

                    # [+]V1.3.0.0 (2014/05/23)
                    $assembly_loaded = $true
                }


                if (($destination_Path | Test-Path) -and (-not $Force))
                {
                    # File Already Exist
                    # [*]V1.1.1.0 (2014/02/27) Change from Write-Verbose into Write-Warning
                    Write-Warning ('[' + $MyInvocation.MyCommand.Name + ']' `
                        + " WARNING: Zip Compression of '$source_Filename' is aborted because '$destination_Path' already exists!")
                    
                    # [-]V1.3.0.0 (2014/05/23)
                    # $aborted = $true

                    # [+]V1.1.0.0 (2014/02/23)
                    $destination_Path = [string]::Empty
                }
                else
                {
                    # Force
                    if ($destination_Path | Test-Path) { $destination_Path | Remove-Item -Force -Recurse }

                    # Zip (by .NET) / [*]V1.0.5.0 (2014/01/17)
                    if ((Get-Item -Path $InputObject) -is [System.IO.DirectoryInfo])
                    {
                        # Directory
                        [System.IO.Compression.ZipFile]::CreateFromDirectory($source_Path, $destination_Path)
                    }
                    else
                    {
                        # File
                        [System.IO.Compression.ZipArchive]$archive = `
                            [System.IO.Compression.ZipFile]::Open($destination_Path, ([System.IO.Compression.ZipArchiveMode]::Create))

                        [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $source_Path, ($source_Path | Split-Path -Leaf))
                    }
                    Write-Verbose ('[' + $MyInvocation.MyCommand.Name + ']' + " '$source_Filename' -> '$destination_Path'")
                    
                    # [-]V1.3.0.0 (2014/05/23)
                    # $completed = $true
                }
            }
            catch  # [+]V1.3.0.0 (2014/05/23)
            {
                if (-not $assembly_loaded)
                {
                    Write-Warning ('[' + $MyInvocation.MyCommand.Name + ']' + " Fail to load '" + ($script:AssemblyName -split ',')[0] + "'!" )
                }
                else
                {
                    # [+]V1.3.1.0 (2014/05/25)
                    Write-Error $_
                }

                $destination_Path = [string]::Empty
            }
            finally
            {
                if ($archive -ne $null) { $archive.Dispose() }
            }
        }


        # Change condition / [*]V1.3.0.0 (2014/05/23)
        if (-not $assembly_loaded)
        {
            # Shell mode
            Write-Verbose ('[' + $MyInvocation.MyCommand.Name + '](Shell)' + " Enter Shell mode...")

            # Decompression of Directory in Shell mode is not supported... / [*]V1.0.5.0 (2014/01/17)
            if ((Get-Item -Path $source_Path) -is [System.IO.DirectoryInfo]) { throw New-Object System.NotSupportedException }

            # Add Zip Header / [*]V1.0.4.0 (2014/01/14)
            $zip_Header = [System.Convert]::ToByte([char]'P'), [System.Convert]::ToByte([char]'K'), [byte[]](0x05, 0x06), ([byte[]]0x00 * 18)
            $zip_Header | Set-Content -Path $destination_Path -Encoding Byte

            ($destination_Path | Get-ChildItem).IsReadOnly = $false
        
            $shell = New-Object -ComObject Shell.Application
            $zip = $shell.NameSpace($destination_Path)

            # Create ZIP File (Shell mode)
            $zip.CopyHere(([System.IO.FileInfo]$source_Path).FullName)
            Write-Verbose ('[' + $MyInvocation.MyCommand.Name + '](Shell)' + " '$source_Filename' -> '$destination_Path'")
        }


        # Output -> Return / [*]V1.3.0.0 (2014/05/23)
        return $destination_Path
    }
}

#####################################################################################################################################################
