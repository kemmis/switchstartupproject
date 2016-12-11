# Configuration File

SwitchStartupProject can be configured with a solution specific configuration file.
It allows to define the available startup project entries in the dropdown and their behavior.

To open the configuration file in Visual Studio, select `Configure...` in the SwitchStartupProject dropdown.
A new configuration file with default values (and some comments) is created if it does not exist yet.

## Location and Filename

The configuration file is located in the same folder as the solution file, it has the same name as the solution file, but uses extension `.sln.startup.json`.

| Example | |
| --- | --- |
| Visual Studio Solution File | `MySolution.sln` |
| SwitchStartupProject Configuration File | `MySolution.sln.startup.json` |

## Format

The configuration file uses [JSON](http://json.org/) syntax with C-style `/* ... */` comments.
See also the [JSON schema](https://bitbucket.org/thirteen/switchstartupproject/src/tip/SwitchStartupProject/Configuration/ConfigurationSchema.json) that is used to validate the configuration file.

## Example

```
#!json
{
    "Version": 3,
    "ListAllProjects": false,
    "MultiProjectConfigurations": {
        "A + B (CLA + WD)": {
            "Projects": {
                "MyProjectA": {},
                "MyProjectB": {
                    "CommandLineArguments": "1234",
                    "WorkingDirectory": "%USERPROFILE%\\test\"
                }
            }
        },
        "A + B": {
            "Projects": {
                "MyProjectA": {},
                "MyProjectB": {
                    "CommandLineArguments": "",
                    "WorkingDirectory": ""
                }
            }
        },
        "D": {
            "Projects": {
                "MyProjectD": {}
            }
        }
    }
}
```

## Configuration Properties

### Version
The version of the configuration file format. Must be `3`.

### ListAllProjects
Can be either `true` or `false`.

If set to `true`, SwitchStartupProject creates an item in the dropdown for each project in the solution, which allows each project to be activated individually as startup project.

### MultiProjectConfigurations
A dictionary of named startup configurations, each consisting of one or multiple startup projects with optional parameters like command line arguments and working directory.

Example of a startup configuration:
```
        "A + B (CLA) + C": {                                    /*  Configuration name (appears in the dropdown)  */
            "Projects": {
                "MyProjectA": {},                               /*  Starting project A  */
                "MyProjectB": {                                 /*  and project B ...   */
                    "CommandLineArguments": "1234",             /*  ... with command line arguments "1234"  */
                    "WorkingDirectory": "%USERPROFILE%\\test"   /*  ... with working directory %USERPROFILE%\test  */
                },
                "Path\\To\\ProjectC.csproj": {}                 /*  and project C (specified by path)  */
            }
        }
```
Startup projects can be specified in two ways:

* Either by project name as shown in the solution explorer of Visual Studio.
* Or by the path to the project file, relative to the solution file. This allows to unambiguously refer to projects with same name. Note that backslashes have to be escaped (duplicated) in JSON.

SwitchStartupProject creates an item in the dropdown for each startup configuration, allowing the corresponding startup projects (and command line arguments) to be activated.

> Note:
> If a startup project specifies a parameter (like command line arguments or a working directory), that parameter is (persistently) set when the configuration gets activated. If a startup project does not specify a parameter, the existing parameter value won't change when the configuration gets activated. To clear a parameter set by another configuration, specify the parameter with an empty string `""` value.

## Default Values

If SwitchStartupProject finds no configuration file, it uses the following defaults:

```
#!json
{
    "Version": 3,
    "ListAllProjects": true,
    "MultiProjectConfigurations": {}
}
```

## Configuration Updates

SwitchStartupProject reads and applies the configuration file of a solution

* when the solution is opened
* whenever the configuration file is changed while the solution is open

## Older Versions

Up to SwitchStartupProject version `3.0` the configuration file was named `MySolution.startup.suo`.
