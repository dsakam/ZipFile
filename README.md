ZipFile
=======

Zip File Compression / Decompression Module for PowerShell Version **1.0.4.0**


概要
----

Zip ファイルの作成と解凍を行うための PowerShell モジュールです。



履歴
----

**V1.0.4.0** (2014/01/14)  
New-ZipFile コマンドレットの Zip ヘッダーを書き込むときの型を object 型から byte 配列に変更。

**V1.0.3.0** (2014/01/10)  
Expand-ZipFile コマンドレットの InputObject パラメーターにフォルダーが指定された時の例外処理を追加しました。

**V1.0.2.0** (2014/01/06)  
Expand-ZipFile コマンドレットの InputObject パラメーターに空のファイルを指定した時の例外を System.IO.InvalidDataException から System.IO.FileFormatException に変更しました。

**V1.0.1.0** (2014/01/06)  
Expand-ZipFile コマンドレットの InputObject パラメーターのバリデーション (ValidateScript) に処理を追加。  
また、Copyright に 2014 年を追加。

**V1.0.0.3** (2013/12/28)  
README.md の小変更。

**V1.0.0.2** (2013/12/17)  
コメントや README.md 等の小変更。

**V1.0.0.1** (2013/12/13)  
ファイル削除の警告メッセージを Wirte-Warning から Write-Verbose に変更。

**V1.0.0.0**  (2013/12/12)  
1st Edition


Expand-ZipFile
--------------

Zip ファイルのファイル名から拡張子 (.zip) を除いたフォルダを作成し、その中に Zip ファイルを解凍します。

**Expand-ZipFile** コマンドレットは、まず System.IO.Compression.FileSystem.dll のロードを試みます。  
System.IO.Compression.FileSystem.dll のロードに成功すると、System.IO.Compression.ZipFile クラスの ExtractToDirectory メソッドを使って Zip ファイルを解凍します。

System.IO.Compression.FileSystem.dll のロードに失敗すると、シェルモードで Zip ファイルを解凍します。  
シェルモードでは、System.Shell.Folder.CopyHere メソッドを使って Zip ファイルを解凍します。


New-ZipFile
-----------

指定されたフォルダーまたはファイルから、フォルダー名またはファイル名に拡張子 .zip を付加した Zip ファイルを作成します。

**New-ZipFile** コマンドレットは、まず System.IO.Compression.FileSystem.dll のロードを試みます。  
System.IO.Compression.FileSystem.dll のロードに成功すると、圧縮対象がフォルダーの場合は System.IO.Compression.ZipFile クラスの CreateFromDirectory メソッドを使って Zip ファイルを作成します。  
圧縮対象がファイルの場合は、System.IO.Compression.ZipArchive クラス、および System.IO.Compression.ZipFile クラスと System.IO.Compression.ZipFileExtensions クラスの CreateEntryFromFile 拡張メソッドを使って Zip ファイルを作成します。

System.IO.Compression.FileSystem.dll のロードに失敗すると、シェルモードで Zip ファイルを作成します。  
シェルモードでは、System.Shell.Folder.CopyHere メソッドを使って Zip ファイルを作成します。  
シエルモードでは、フォルダーの圧縮をサポートしていません。
