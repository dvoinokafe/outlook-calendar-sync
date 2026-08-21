# Security

## Reporting a vulnerability

Please report security issues privately through GitHub's security advisory mechanism if enabled for this repository. Do not include company calendar data, credentials, email addresses, or confidential appointment content in a public issue.

## Data handling

The synchronizer reads calendar items through the Classic Outlook COM object model and writes mirrored appointments into another Outlook calendar in the same local Outlook profile.

Calendar content may be sensitive. The user is responsible for ensuring that cross-company copying is permitted by both organizations.

## Credentials

This project does not request Microsoft Graph client secrets, OAuth application credentials, or tenant-admin consent. Authentication is inherited from the Outlook profile already signed into Windows.

Do not add passwords, tokens, exported Outlook profiles, or company-specific secrets to this repository.

## Local files

The project stores:

- synchronization mappings in `%LOCALAPPDATA%\OutlookCalendarSync\sync.db`
- logs in `%LOCALAPPDATA%\OutlookCalendarSync\sync.log`

Protect these files according to your organization's requirements.

## Scope

This software is intended for legitimate synchronization of calendars the user is authorized to access. It does not attempt to bypass Outlook, Microsoft 365, Windows, or employer security controls.
