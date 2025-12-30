# outlook-web

[![CI](https://github.com/Osso/outlook-web/actions/workflows/ci.yml/badge.svg)](https://github.com/Osso/outlook-web/actions/workflows/ci.yml)
[![GitHub release](https://img.shields.io/github/v/release/Osso/outlook-web)](https://github.com/Osso/outlook-web/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

CLI for Outlook Web via Chrome DevTools Protocol (browser automation).

## Installation

```bash
cargo install --path .
```

## Setup

Launch Chrome/Chromium with remote debugging:
```bash
chromium --remote-debugging-port=9222
```

## Usage

```bash
outlook-web list              # List inbox messages
outlook-web list-spam         # List junk folder
outlook-web read <id>         # Read a specific message
outlook-web archive <id>      # Archive message
outlook-web spam <id>         # Mark as spam
outlook-web label <id> <cat>  # Add category
outlook-web delete <id>       # Delete message
outlook-web test              # Test browser connection
```

## License

MIT
