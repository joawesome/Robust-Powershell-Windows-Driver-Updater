# Robust-Powershell-Windows-Driver-Updater
This is a more "error free" approach for using windows updates to install missing drivers on fresh new windows installs.

This was inspired by @[SapphSky](https://github.com/SapphSky)'s driver update script in his win.refurb.sh repository, as we use it for our company to make driver installation more efficent. 

## How it works

This uses the [PSWindowsUpdate](https://www.powershellgallery.com/packages/pswindowsupdate) module to install device drivers through windows updates. Due to the nature of windows updates errors can occur when installing multiple drivers at once. This is mostly due to a caching issue. To overcome this, this script reruns the module if the error occurs. This minimizes user interaction in the case of an error so that the user do not have to rerun the module.

