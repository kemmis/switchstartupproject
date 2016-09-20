# Release notes

## Version 3.1
* Less smart, less GUI, more stability, more power for the user
  * Removed smart mode: It was not smart enough and often failed with newer project types. (issue #30)
  * Removed MRU mode: It was not really useful and made it hard to share configurations. (issue #29)
  * Removed GUI: Visual Studio and the GUI frequently caused problems when storing the configurations. (issues #26, #28)
  * Configuration file is now read only. (issue #31)
  * Immediately apply configuration file changes to currently active startup projects. (issue #32)
* New configuration file extension: <SolutionName>.startup.json
* Reorganize readme, license and release notes files

## Version 3.0
* Add support for the upcoming Visual Studio "15"
* Continuous Integration builds on AppVeyor (issue #34)

## Version 2.8
* Add support for F# Exe/WinExe projects in Smart Mode (thanks to TeaDrivenDev)
* Add support for command line arguments (issue #21)
* Add button to clone multi-project startup configuration (issue #18)
* Store single project mode for each solution (issue #22)
* Fix issue #11: Select the current startup project in the dropdown when a project has been (re-)loaded
* Fix issue #20: Select last used multi-project configuration in dropdown box when re-opening a solution
* Fix issue #23: Make sure multi-project configurations of last opened solution are not deleted when creating a new project/solution
* Fix issue #24: Handle renaming of projects

## Version 2.7
* Fix issue #17: Prevent exceptions when creating a new website.
* Re-enable support for Visual Studio 2014 (see issue #13).

## Version 2.6
* Fix issue #16: Fix MRU mode: Most recently used projects should only show projects that are in the solution
* Fix issue #14: Add support for visual fortran projects (and other project types that don't implement IVsAggregatableProject)

## Version 2.5
* Fix issue #12: Fix web site projects support

## Version 2.4
* Track configuration file and reload settings upon change
* Save configuration file whenever solution gets saved
* Smart mode: Improve detection of startable projects
* Support for Visual Studio 2014
* Migrate source code to Visual Studio 2013

## Version 2.3
* Fix issue #10: Improve behavior with projects that are not in dropdown list
* Fix issue #9: Support Azure projects
* Fix issue #5: Support VS 2013 projects in smart mode

## Version 2.2
* Fix issue #8: Fix multi project configurations in VS 2010

## Version 2.1
* Fix issue #6: Enable dropdown after installation and restart
* Fix issue #7: Support VS 2012 projects loading in background

## Version 2.0
* Support for multiple startup projects.