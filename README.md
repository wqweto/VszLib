## VszLib

7-zip SDK supports several compression methods and can produce and read 7z, zip, gzip, tar, bzip2 and other archives. This is a VB6 helper component that makes using original 7z.dll in your VB6 projects possible. 

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

### More compression

    With New cVszArchive
        .AddFile "your_file"
        .AddFile "folder\another_file", "folder\name_in_archive"
        '--- fast compression (-mx3)
        .Parameter("x") = 3 
        '--- LZMA2 method support multi-core compression
        .Parameter("0") = "LZMA2"
        .Parameter("mt") = "on"
        .Password = "secret"
        '--- split in 10MB volumes
        .VolumeSize = 10# * 1024 * 1024
        .CompressArchive "test_lzma2.7z"
    End With
