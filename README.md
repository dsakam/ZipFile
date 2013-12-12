ZipFile モジュール
==================

Zip File Compression and Decompression Module for PowerShell


Expand-ZipFile コマンドレット
-----------------------------

Zip ファイルのファイル名から拡張子 (.zip) を除いたフォルダを作成し、その中に Zip ファイルを解凍します。

__Expand-ZipFile__ コマンドレットは、まず System.IO.Compression.FileSystem.dll のロードを試みます。  
System.IO.Compression.FileSystem.dll のロードに成功すると、System.IO.Compression.ZipFile クラスの ExtractToDirectory メソッドを使って Zip ファイルを解凍します。

Force パラメーターが指定されていないときは、System.IO.Compression.ZipArchive クラスを使って Zip ファイル内のエントリーをチェックします。  
Zip ファイルに含まれるファイルの数が多い場合はパフォーマンスに影響する可能性があるので、そのようなときは Force オプションを指定することを検討してください。

System.IO.Compression.FileSystem.dll のロードに失敗すると、シェルモードで Zip ファイルを解凍します。  
シェルモードでは、CopyHere メソッドを使って Zip ファイルを解凍します。


New-ZipFile コマンドレット
--------------------------

指定されたフォルダーまたはファイルから、フォルダー名またはファイル名に拡張子 .zip を付加した Zip ファイルを作成します。

__New-ZipFile__ コマンドレットは、まず System.IO.Compression.FileSystem.dll のロードを試みます。  
System.IO.Compression.FileSystem.dll のロードに成功すると、圧縮対象がフォルダーの場合は System.IO.Compression.ZipFile クラスの CreateFromDirectory メソッドを使って Zip ファイルを作成します。  
圧縮対象がファイルの場合は、System.IO.Compression.ZipArchive クラス、および System.IO.Compression.ZipFile クラスとSystem.IO.Compression.ZipFileExtensions クラスの CreateEntryFromFile 拡張メソッドを使って Zip ファイルを作成します。

System.IO.Compression.FileSystem.dll のロードに失敗すると、シェルモードで Zip ファイルを作成します。  
シェルモードでは、CopyHere メソッドを使って Zip ファイルを作成します。  
シエルモードでは、フォルダーの圧縮をサポートしていません。  
