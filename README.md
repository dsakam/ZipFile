ZipFile
=======

Zip File Compression / Decompression Module for PowerShell Version **1.3.1.0**


概要
----

Zip ファイルの作成と解凍を行うための PowerShell モジュールです。



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


履歴
----

**V1.3.1.0** (2014/05/25)  
Expand-ZipFile / New-ZipFile 両コマンドレットのヘルプコンテンツを修正。  
New-ZipFile コマンドレットの出力先が不正かどうかのチェックを追加。  

**V1.3.0.0** (2014/05/23)  
Expand-ZipFile / New-ZipFile 両コマンドレットの主に例外処理を大幅に変更。  
New-ZipFile コマンドレットの出力先が入力元と同じ場合はエラーになるように変更。  

**V1.2.0.0** (2014/05/15)  
Expand-ZipFile コマンドレットの無意味なパラメーターチェックにより OutOfMemoryException になる不具合を修正。  

**V1.1.2.0** (2014/02/27)  
一部の引用符（シングルクォーテーションとダブルクォーテーション）の変更。  

**V1.1.1.0** (2014/02/27)  
コメントのスタイルを少し変更しました。  
New-ZipFile コマンドレットで、既にファイルが存在するときのコンソール出力を冗長出力 (Write-Verbose) から、警告 (Write-Warning) に変更しました。

**V1.1.0.0** (2014/02/23)  
引用符（シングルクォーテーションとダブルクォーテーション）の扱いを変更しました。  
冗長コンソール出力 (Write-Verbose) の内容を変更しました。  
New-ZipFile コマンドレットで、既にファイルが存在するために Zip ファイルを作成しない場合の出力を [string]::Empty に変更しました。

**V1.0.5.0** (2014/01/17)  
型チェックの方法を変更しました。

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
