# Office-2021-Download-and-Activate
## How to download oficial and complete Microsoft Office 2021 and activate it using KMS client key


## Download and install Office 2021
First, make sure that the operating system version you are running is Windows 10 or later before you get started. There is no way to install Office 2021 on Windows 8.1 or earlier.
Just click [here](https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2021Retail.img)  to get the official install image from Microsoft. The size of the IMG file is 4.2 GB. Double click on the file to mount it to a virtual drive on your PC once the download is complete. Then follow the instructions of the Setup Wizard to install Office 2021 on your Windows.
This might take a while, please wait.


## Activate Office 2021 for FREE using KMS client key

Open cmd program with administrator rights.
First, you need to open cmd in the admin mode, then run all commands below one by one.
If you install your Office in the ProgramFiles folder, the Office directory depends on the architecture of your OS. If you are not sure of this issue, just run both of the commands above. One of them will be not executed and an error message will be printed on the screen.
```
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16

cd /d %ProgramFiles%\Microsoft Office\Office16
```

This step is required. You can not install the KMS client product key of Office without a volume license.
```
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
```


## Activate your Office using the KMS key.
Make sure your device is connected to the internet, then run the following commands.
```
cscript ospp.vbs /setprt:1688

cscript ospp.vbs /unpkey:6F7TH >nul

cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH

cscript ospp.vbs /sethst:e8.us.to

cscript ospp.vbs /act
```

If you see the error 0xC004F074, it means that your internet connection is unstable or the server is busy. Please make sure your device is online and try the command “act” again until you succeed.
