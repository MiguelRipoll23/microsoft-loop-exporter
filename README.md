# Microsoft Loop Exporter

A Node.js tool to export all pages in a Microsoft Loop workspace to PDF files
automatically.

## Demo

TODO

## Installation

You can install this package globally using npm:

```bash
npm install -g microsoft-loop-exporter
```

## Usage

> [!IMPORTANT]
> Close your Chrome browser before executing the command.

Run the following command:

```bash
loop "Your Workspace Name"
```

Replace "Your Workspace Name" with the exact name of your Loop workspace as it
appears in the Loop interface.

### Troubleshooting

- If you get a "Chrome not found" error, verify that Chrome is installed in the
  default location for your OS
- If workspace items are not found, ensure:
  - You're logged into Microsoft Loop
  - The workspace name is exactly as it appears in Loop
  - You have proper access to the workspace
