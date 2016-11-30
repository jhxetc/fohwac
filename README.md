# gc
the ghetto wow addon update client

Can be run interactively or in the background to download and install WoW addons using curseforge and wowace project URLs.

Simply call .\gc.ps1 to run interactively or call it with the -background paramater to update all addons. To update a single addon, call with '-background -addon <number>' where <number> is the number of the addon you want to update.

All configuration information is stored within a json file in the same directory as the script. It will automatically generate upon first run and prompt you to fix the WoW addon path if it is different from the default WoW path. You can edit the json file directly as long as you keep proper json formatting. You can copy json files from others, or keep multiple json files for different addon sets. 

examples:

.\gc.ps1 : run interactively
.\gc.ps1 -background : download and install all addons.
.\gs.ps1 -background -addon 1 : download and install addon number 1. Addon number can be obtained by listing addons in interactive mode, or by counting addons in the json file where 0 is the first entry.
