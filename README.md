# Welcome to PowerShell AutoSort
###  Explanation
---
This is a **Windows PowerShell Script** made in March 2017 for Eston Manufacturing, Guelph, ON. 

Traditionally, CMM Inspection reports were saved in PDF format. At the end of each shift, CMM operators need to manually count how many in process parts failed for inspection during their shift so that they can provide the information to the management level. Eventually, all the failed PDF reports need to be archived. 

This PowerShell Script was designed to automatically count how many failed inspections reports in the current folder by the name of the files. Then it will temporarily move all the failed in-process reports to a folder called "In-Process-Fail" for future review, a .bat file called "Step2.bat" is also going to be created there. If everything is OK, people can double click "Step2.bat" to send all the failed reports to archive them.  


###  Usage
---

1. Put the AutoSort.ps1 script together with the PDF files you want to sort

2. Setup the root directory where you want to archive the failed results  in line 97 of AutoSort.ps1 file.
```batch
set rootdir=E:\Google Drive\Auto Sort\Demo\Archive_Folder
```
3. Start the PowerShell Script with PowerShell ISE and type in the time range that within how many hours we sort the result files
4. Review all the sorted failed reports in the "In-Process Fail" folder, and click Setp2.bat to send them to the archive folder.

<img src="https://github.com/y5mei/Saved-Pictures/blob/master/SortInstruction.gif" alt="ins" style="zoom:100%;" />

### How to active PowerShell on your computer
---
There are 4 ways you can active PowerShell on your computer:

1. To active PowerShell on you computer, you can do: win+R cmd -> PowerShell -> `Set-ExecutionPolicy RemoteSigned`
2. Or, at the beginning of you script, write down: `Set-ExecutionPolicy RemoteSigned`
3. Or, You can try and set the policy of the process itself: powershell.exe -`ExecutionPolicy bypass`
4. If your domain administrator hasn't forbidden it, you can do this:
`Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser`