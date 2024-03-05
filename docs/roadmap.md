# ish Roadmap

## Pre-Alpha Phase (v0.1.x - v0.2.x)

- Research and Design (v0.1.0)
  - Finalize the command list for the initial release, focusing on common command-line and bash commands.
  - Design the architecture for the ish core, including how it will interact with Excel, VBA, and external processes.
  - Outline the YAML schema for configuration management.
- Core Shell Functionality (v0.2.0)
  - Implement basic shell command execution within the Immediate Window (e.g., ish.cls/clear, ish.echo).
  - Develop a basic package manager framework within ish, capable of managing AddIn dependencies and installations.

## Alpha Phase (v0.3.x - v0.5.x)

- Package Mananger Development (v0.3.0)
  - Implement the hidden Excel sheet for package registration with ListObject getters.
    - May not be embedded in the sheet for update purposes, may move to pure YAML in the future
  - Develope functionalities for installing, updating, and removing packages.
- Configuration Management (v0.4.0)
  - Implement a lightweight YAML parser for reading and writing configuration files.
  - Create a basic user form for modifying the configurations, backed by a configuration manager.
- Common Commands Implementation (v0.5.0)
  - Implement additional common command execution within the Immediate Window (e.g., ish.ls/dir, ish.cd).
  - Test and refine the execution of these commands in diverse Office environments (e.g., Outlook, Word, PowerPoint).

## Beta Phase (v0.6.x - v0.9.x)

- Project Manager AddIn (v0.6.0)
  - Develop the Project Manager as the first official package, utilizing the package manager and configuration manager.
  - Implement project configuration consumption and management within the Project Manager.

- Enhanced Shell Commands (v0.7.0 - v0.9.0)
  - Incrementally add support for more complex command-line and bash commands.
  - Focus on ensuring compatibility across different Windows versions and considering cross-platform possibilities.

- Testing and Documentation (v0.9.0)
  - Conduct extensive testing across all functionalities, focusing on stability, performance, and security.
  - Write comprehensive documentation, including user guides, API documentation, and developer guides.

## Release Candidate (v0.10.x)

- Polishing and Bug (v0.10.0)
  - Address all known bugs and polish user experience based on beta feedback.
  - Finalize all documentation and prepare marketing materials for launch.

## Version 1.0.0 - Official Release

- Launch (v1.0.0)
  - Officially release the ish core, package manager, and Project Manager AddIn.
  - Promote the release through relevant channels to the VBA and Excel communities.

## Post-Release (v1.1.x and beyond)

- Feedback Loop and Iterations (v1.1.0)
  - Collect user feedback and prioritize new features or commands for future updates.
  - Begin development on less common, rare, and obscure commands based on demand.
- Git AddIn Development
  - Start the development of the git AddIn as a separate project, leveraging the ish infrastructure.

This roadmap is designed to be iterative, allowing for continuous feedback and adjustments as needed. The focus on delivering a stable and useful v1.0.0 release should ensure that the project provides immediate value to users, with ample room for growth and expansion in future versions.
