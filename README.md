# FoH WoW Addon Client

Can be run interactively or in batch mode to download and install WoW addons using curseforge and wowace project URLs.

Simply call .\fohwac.ps1 to run interactively or call it with the -batch paramater to update all addons. To update a single addon, call with '-batch -addon <number>' where <number> is the number of the addon you want to update. To import another addon file use the -json parameter followed by either a web address to json output or a local file path. If you have json loaded from the web any addon entries that you add/change/remove will only persist until you close the program. 

All configuration information is stored within 2 json files in the same directory as the script. It will automatically generate both files upon first run and prompt you to fix the WoW addon path if it is different from the default WoW path and also for a current WoW version (currently 7.1.0). You can edit the json files directly as long as you keep proper json formatting. You can copy json files from others, or keep multiple json files for different addon sets.

If you edit json manually: There are 3 properties for each addon entry: AddonName, AddonUrl and AddonRel. The first 2 are self explanatory and the last accepts either a 0 (off) or 1 (on). This property determines whether the client will download and install all versions (alpha/beta/release) or just release versions. For all versions set to 0 and for only release versions set to 1. 

examples:

.\fohwac.ps1 : run interactively
.\fohwac.ps1 -batch : download and install all addons.
.\fohwac.ps1 -batch -addon 1 : download and install addon number 1. Addon number can be obtained by listing addons in interactive mode, or by counting addons in the json file where 0 is the first entry.
.\fohwac.ps1 -batch -json http://www.myguild.com/addons.json : download addons from the json output at this url. Good for raid leaders who want to require a standard set of addons.

You can use most combinations of parameters together.
