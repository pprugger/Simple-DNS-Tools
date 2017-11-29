# Simple-DNS-Tools
A few simple scripts to change DNS NIC settings on windows

To use the change-dns scripts just change the computer name in show-dns-settings.vbs to your computer name and execute it.
It will show you all your configured NICs.
Remember the index number and change ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE Index = 1") in the change-dns scripts to your desired Index number.

If you want to change the DNS settings of all NICs simultaneously change ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE Index = 1") to ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True") like seen in the show-dns-settings.vbs script.

The reset and set dns admin scripts are powershell scripts. They can be used if you have no admin rights with your current user and you have the credentials for an admin user.
Just change the paths in the ps1 files to your desired script path.

Dont forget to change the computer name in all scripts.
If you want the scripts to generate output messages just uncomment the WScript.Echo lines.



