## VszLib 
(Vsz = Vb Seven Zip)

[7-zip SDK](http://www.7-zip.org/sdk.html) supports several compression methods that can produce and read 7z, zip, gzip, tar, bzip2 and other archives. This is a VB6 helper component that makes using original `7z.dll` in your VB6 projects possible. 

### Simple compress

    With New cVszArchive
        .AddFile "your_file"
        .CompressArchive "test.7z"
    End With

### Extract

    With New cVszArchive
        .OpenArchive "test.7z"
        .Extract "extract_folder"
    End With

### Create zip file

    With New cVszArchive
        .AddFile "your_file"
        .CompressArchive "archive.zip"
    End With

### More options

    With New cVszArchive
        .AddFile "your_file"
        .AddFile "folder\another_file", "folder\name_in_archive"
        '--- fast compression (-mx3)
        .Parameter("x") = 3 
        '--- LZMA2 method supports multi-core compression
        .Parameter("0") = "LZMA2"
        .Parameter("mt") = "on"
        .Password = "secret"
        '--- split in 10MB volumes
        .VolumeSize = 10# * 1024 * 1024
        .CompressArchive "test_lzma2.7z"
    End With
	
    With New cVszArchive
        .OpenArchive "source.7z"
        .Extract "bin_folder", "*.exe"
    End With

### Using source

Before opening `Src\VszLib.vbp` register `Src\SevenZip.tlb` with `regtlib.exe`, VB6 IDE or your favorite typelib registration tool. You don't need to redistribute `SevenZip.tlb`, it's only needed in development environment.

### Using component

Register `Bin\VszLib.dll` with `regsvr32.exe` (or VB6 IDE) and add a reference (Project | References...) in your project to *7-zip VB6 Helper 1.0*. You only need to redistribute `Bin\VszLib.dll` with your application, `pdb` files are needed only for debugging purposes. Note that `Bin\VszLib.dll` is compiled with line numbers in the source so that error logging can produce more useful traces.

### API

The only publicly accessible class in the library is `cVszArchive`. Here is a short description of the methods, properties and events in order of relevance (kind of).

#### `Init([DllFile As String]) As Boolean`

Optionally used to indicate `7z.dll` location. If `DllFile` is empty first `7z.dll` is loaded from `VszLib.dll` folder, then registry is inspected for 7-zip setup folder, finally lightweight `7za.dll` is attempt loaded from `VszLib.dll` folder. Best practice is to place `7z.dll` next to `VszLib.dll` and to not call `Init` explicitly from client code. Extract/compress operation will call it if needed.

Note that `7za.dll` (from [7-zip extras](http://sourceforge.net/projects/sevenzip/files/7-Zip/9.22/7z922_extra.7z/download)) can be used to compress/extract only 7z archives (no zip support). The even smaller `7zxa.dll` (172KB) can be used to only extract 7z archives.

#### `OpenArchive(ArchiveFile As String) As Boolean`

Opens the archive using archive file extension to guess decompressor type. Currently supported file extension for update: 7z, zip, tar, bz2, gz, xz, wim. Other formats supported: rar, cab, chm, iso, msi, hfs, iso, arj, cpio, deb, dmg, fat, flv, lzh, lzma, lzma86, mbr, ntfs, exe, pmd, rpm, 001, swf, vhd, xar, z. Populates `FileCount` and `FileInfo` properties. 

#### `Extract(TargetFolder As String, [Filter]) As Boolean`

Extracts files to `TargetFolder` from a previously opened archive file. Optional `Filter` can specify which file entries to extract by exact match (`document.txt`), a filename mask (`*.exe`) or an array of `FileCount` booleans, each index indicating whether to decompress the file with same index (array entry set to `True`) or to skip it (array entry set to `False`). Raises `Progress` event to indicate progress and to allow cancellation of the extraction.

#### `Property FileCount As Long` (read-only)

Returns number of file entries in the archive

#### `Property FileInfo(FileIdx As Long)` (read-only)

Returns an array with information about a file entry. Array indexes are: 0 - file name, 1 - attributes, 2 - size, 3 - bool if encrypted, 4 - CRC, 5 - file comment, 6 - creation time, 7 - last access time, 8 - last write time. Some of the entries can be `Empty` if not supported by the current archive format.

#### `AddFile(File As String, [Name As String], [Comment As String]) As Boolean`

Adds a file to archive. `File` must be an (absolute) path to an existing file. Optional `Name` can specify name and relative folder in the archive the entry is going to be stored to. If `Name` not specified, filename portion of `File` is used as name in the root folder of the archive. `Comment` is optional and (probably) not supported by all compressors.

#### `CompressArchive(ArchiveFile As String) As Boolean`

Creates an archive using previously added files. Compressor type is guessed by archive file extension. Raises `Progress` event to indicate progress and to allow cancellation of the compression.

#### `Property Parameter(ParamName As String) As Variant` (read/write)

Specifies custom compression parameters. These correspond to `-m` switch of command line `7z.exe`. 
Setting compression level switch `-mx3` or `-mx=3` is translated to `Parameter("x") = 3`.
Setting 7z compression method `-m0=LZMA2` is translated to `Paramter("0") = "LZMA2"`.
Setting multi-threading switch `-mmt=on` or `-mmt=3` is translated to `Paramter("mt") = "on"` or `Paramter("mt") = 3`.
Setting encrypt headers switch `-mhe=on` is translated to `Paramter("he") = "on"`. See [more examples of -m switch](http://www.dotnetperls.com/7-zip-examples) for additional info.

#### `Property Password As String` (read/write)

Gets/sets password used during extraction/compression. If incorrect archive password is using, decompressor raises *Data Error. Wrong password?* error through `Error` event.

#### `Property VolumeSize As Double` (read/write)

Gets/sets volume size in bytes for split volumes to be created during compression. Can be used with 7z archives only. For other formats splits output archive in `VolumeSize` sized chunks which is not supported model by native decompressors.

#### `Property FormatCount As Long` (read-only)

Gets number of formats supported by `7z.dll` which has been loaded on `Init`.

#### `Property FormatInfo(FormatIdx As Long) As Variant` (read-only)

Returns an array with information about compression format. Array indexes are: 0 - name, 1 - class ID, 2 - file extension(s), 3 - additional extensions (if any), 4 - bool if update supported, 5 - bool if keeps names, 6 - byte array with archive start signature, 7 - byte array with archive finish signature.

#### `Property LastError As String` (read-only)

Gets last error that occurred during last operation. Returns empty straint if no error occurred.

#### `Event Progress(FileIdx As Long, Current As Double, Total As Double, Cancel As Boolean)`

Raised when new information about current operation progress is available and to give the user a chance to cancel operation. `FileIdx` is the index of the current file being extracted/compressed. This index can be used with `FileInfo` property. `Current` and `Total` parameters can be using to calculate percentage completed. `Cancel` can be set to abort current operation. 

#### `Event Error(Description As String, Source As String, Cancel As Boolean)`

Raised when an unexpected condition occurs during current operation. `Description` contains *User cancelled* when `Cancel` gets set in `Progress` event. `Description` contains *Data Error. Wrong password?* if wrong password is supplied for extraction of encrypted archives. Can be raised when input/output files are inaccessible (permissions, network outages). Setting `Cancel` indicates whether to finish operation or abort it immediately.

#### `Event NewVolume(FileName As String)`

Raised when creating multi-volume archives to indicate new file names that are created during compression. Useful if volumes are to be deleted on error or user cancellation.
