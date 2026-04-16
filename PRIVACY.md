# Privacy Policy — Azure DevOps Wiki Clone Page

**Last updated: 2026-04-16**

This Chrome extension ("the extension") adds a "Clone" option to the right-click context menu on Azure DevOps wiki pages, allowing the user to duplicate a wiki page within the same wiki.

## What data the extension accesses

When the user activates the Clone action, the extension reads the following from the user's currently-open Azure DevOps wiki, via the official Azure DevOps REST API:

- The selected wiki page's path and content (text/markdown).
- The wiki's page-tree structure (page paths) so the user can choose a destination.

## How the data is used

The data is used solely to perform the requested clone action, which writes a new wiki page back to the same Azure DevOps wiki via the same REST API.

## Data storage and transmission

- The extension does not store any user data.
- The extension does not transmit any user data to any third-party server.
- All API requests are made from the user's browser directly to `dev.azure.com`, authenticated by the user's existing browser session cookie. No credentials are read, stored, or transmitted by the extension.

## Tracking and analytics

The extension contains no analytics, telemetry, advertising, or tracking of any kind.

## Third parties

No data is shared with any third party.

## Contact

For questions about this policy, contact the developer via the GitHub repository: https://github.com/Br3zzly/azure-devops-clone-wiki-page
