# Rubric Builder

Rubric Builder is a local web app for building Demo, Pitch, and POC validation rubrics and exporting them to Excel.

## Quick Start

Download the project, unzip it, then double-click the starter for your computer:

- Mac: `Start Mac.command`
- Windows: `Start Windows.bat`
- Linux: `Start Linux.sh`

The starter will:

- Check that Node.js is available.
- Install dependencies on the first run.
- Start Rubric Builder locally.
- Open the app in your browser at `http://127.0.0.1:4173/`.

## Requirement

Node.js 20 or newer is required. If the starter says Node is missing, install the LTS version from:

https://nodejs.org/

## Current Packaging Note

The Excel export currently depends on `@oai/artifact-tool`, which is private in the Codex runtime. If `npm install` cannot download that package on another machine, the next step is to replace the export dependency with a public package or ship a bundled desktop app.
