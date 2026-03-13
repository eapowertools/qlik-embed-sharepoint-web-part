# qlik-embed-web-part

SharePoint Framework web part for embedding Qlik content in Microsoft 365 pages.

## What It Supports

- Qlik apps
- Qlik sheets
- Qlik charts
- Qlik assistants
- Searchable property-pane selectors for apps, sheets, charts, and assistants
- Fixed and custom embed heights

## Stack

- SharePoint Framework 1.21.1
- TypeScript 5.3
- Qlik API SDK
- Custom SharePoint property-pane controls built with React 17

## Prerequisites

- Node.js `22.14.x`
- A Microsoft 365 tenant with SharePoint Framework development enabled
- A Qlik Cloud tenant host in the format `tenant.region.qlikcloud.com`
- A Qlik OAuth client ID registered with the SharePoint page URL you will use as the redirect URI

## Local Development

```bash
npm install
npm run build
gulp serve
```

## Configuration

In the web part property pane, provide:

1. Tenant host
2. Client ID
3. Embed type
4. App, sheet, chart, or assistant selection
5. Embed height

When tenant and client settings are valid, the web part loads available Qlik resources directly into the property pane so editors can choose from live data instead of pasting IDs manually.

## Notes

- App, sheet, and chart selections are loaded dynamically from the configured tenant.
- Assistant embeds support both legacy and agentic assistant modes.
- Sheet embeds can use classic or analytics UI.
- Chart embeds can use classic or analytics UI.

## Validation

```bash
npm run build
npm test
```
