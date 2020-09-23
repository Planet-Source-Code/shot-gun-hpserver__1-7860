Attribute VB_Name = "Drives"
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" _
        (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
        (ByVal nDrive As String) As Long
       Public Const DRIVE_REMOVABLE = 2
       Public Const DRIVE_FIXED = 3
       Public Const DRIVE_REMOTE = 4
       Public Const DRIVE_CDROM = 5
       Public Const DRIVE_RAMDISK = 6
       Declare Function GetDiskFreeSpace_FAT16 Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
       Dim SectorsPerCluster&, BytesPerSector&, NumberOfFreeClusters&, TotalNumberOfClusters&
Global allDrives As String
Global currDrive As String
Global drvType As String
Global ad As Boolean
Global bd As Boolean
Global cd As Boolean
Global dd As Boolean
Global ed As Boolean
Global fd As Boolean
Global gd As Boolean
Global hd As Boolean
Global id As Boolean
Global jd As Boolean
Global kd As Boolean
