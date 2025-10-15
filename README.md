# Robust-Powershell-Windows-Driver-Updater
This is a more "error free" approach for using windows updates to install missing drivers on fresh new windows installs.

This was inspired by @[SapphSky](https://github.com/SapphSky)'s driver update script in his win.refurb.sh repository, as we use it for our company to make driver installation more efficent. 

## How it works

This uses the [PSWindowsUpdate](https://www.powershellgallery.com/packages/pswindowsupdate) module to install device drivers through windows updates. Due to the nature of windows updates errors can occur when installing multiple drivers at once. This is mostly due to a caching issue. To overcome this, this script reruns the module if the error occurs. This minimizes user interaction in the case of an error so that the user do not have to rerun the module.

## How to use

* Make sure that the computer that this script is running on has internet access
* Open command prompt or powershell as Administrator (or jut press shift + f10 in windows out of box experience)
* Input the following command:
  ''' powershell -ExecutionPolicy Bypass -Command “irm https://raw.githubusercontent.com/joawesome/Robust-Powershell-Windows-Driver-Updater/main/DriverUpdateMain.ps1 | iex” '''
  or if you have a scanner attached to the computer, you can scan this QR code: ![QR Code](qr_command_image.png)

* let the update run and reboot if needed

## Known issues
* Sometimes a whole bunch of updates will fail. Re-run script to get all mising/ failed updates.


### To-dos for future revisions
* Detect failures and retry installation
* Allow an option that reboots computer automatically instead of waiting for user najn'
