# VBish

(I)mmediate Window (SH)ell for VBA

## About ish

The Immediate Window SHell (ish) project aims to transform the Immediate Window in the Visual Basic for Applications (VBA) development environment into a powerful shell-like terminal. This initiative seeks to enhance the productivity and capabilities of VBA developers by enabling the execution of shell commands, script automation, and streamlined project manamement directly within the VBA Editor.

Version 0.1.0 marks the initial release of ish, focusing on laying the groundwork for this ambitius project by introducing core functionalities and the foundation for future enhancements.

## Features

- Shell Command Execution: Execute basic shell commands and scripts directly from the Immediate Window.
- Initial Package Management Framework: A rudimentary systeme to manage the installation and updating of ish add-ins and dependencies.

## Getting Started

### Prerequisites

- Windows 10+ 64-bit 
- 64-bit Microsoft Office
- Basic understanding of VBA and the VBA Editor environment

### Installation

1. Download or build the AddIn, ish.xlam. (Currently .xlsm != .xlam)
  - To download: https://github.com/AttackOnTyler/VBish/raw/main/ish.xlsm
  - To build: clone the repo and place the code files in ./VBish to a standalone AddIn (.xlam)
2. Add the AddIn as a reference to the current project or globally
  - Add Ref: In VBE -> Tools -> References -> Browse -> ish.xlsm
  - Add Global: In App -> File -> Options -> Add-ins -> Manage: Excel Add-ins: Go... -> Browse -> ish.xlsm

### Usage

To start ish, simply type into the Immediate Window `? ish.start[()]`. This should clear the screen, followed by the ish banner, the current working directory of the host workbook, and a new `? ` line.

To execute a command, simply type it into the Immediate Window prefixed with the `? ish.` followed by the command name and parameters. For example:

```vb
? ish.echo("Hello, World!")
```

## Roadmap

For the upcoming releases, we plan to expand the capabilities of ish significantly, including:

- Enhancing the package manager to support a wider rance of add-ins and dependencies.
- Introducing a broader set of shell commands and utilities.
- Developing a user-friendly configuration management interface.
- Implementing the Project Mananger as the first official package.
- Implementing Git wrapper as the second official package.

For more details, see the [project roadmap](docs/roadmap.md).

## Contributing

We welcome contributions from the community! If you're interested in helping to shape the future of ish, please read our [contributing guidelines](docs/contributing.md) for more information on submitting pull requests.

## Support and Feedback

For support request or to provide feedback, please open an issue on the GitHub repository. We're keen to hear from the community and improve ish with your feedback and suggestions.

## License

ish is released under the MIT License. See the [LICENSE](LICENSE) file for more details.

## Acknowledgments

- Thanks to the VBA and developer community for inspiration and support.
- Thanks to all contributors who have helped bring ish to life.
- Special thanks to my girlfriend who has listened to more than enough incoherant ramblings about making the developer experience for VBA better by implementing XYZ ideas I come up with.
