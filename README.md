# Cleanup-TempFolder
## Description
Recursively scans a user-definable temp folder for files at least as old as a specified age, analyzes the folders to determine which 
ones to recycle, and then sends the folders followed by the files to the recycle bin.

In determining which folders to recycle, the script only marks ones that are empty or whose contents -- including all subfolders 
and their files -- are at least as old as the specified age. It only marks the highest-level subfolders that meet those criteria, 
ensuring that the minimum number of folders are sent to the recycle bin.

## Screenshots
WinDirStat showing the number of files and subdirs in the %TEMP% folder prior to running Cleanup-TempFolder:
![WinDirStat - Before](https://github.com/FuLoRi/Cleanup-TempFolder/assets/122180424/a9506780-a964-4f2b-9e88-e11effa25419)

A PowerShell window showing the script running interactively and the user proceeding with default parameters:
![Script - Recycle phase](https://github.com/FuLoRi/Cleanup-TempFolder/assets/122180424/1d1ac25a-bc1c-402b-9031-7a7c1f30270f)

WinDirStat showing the number of files and subdirs in the %TEMP% folder after running Cleanup-TempFolder:
![WinDirStat - After](https://github.com/FuLoRi/Cleanup-TempFolder/assets/122180424/ed9b7fe5-f414-4ff8-bb2f-18f6cc7e1871)

## Known Issues
Performance is slower than desired due to the recursive nature of the script, and is further slowed when Verbose Output and/or Debug Mode are enabled.

