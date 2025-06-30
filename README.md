# Outlook Plugin

Synergia File Saver is a React-based Outlook add-in that allows users to save e-mail attachments to SharePoint and extract contact information directly from Outlook. The project is based on the Office Add-in Task Pane template and uses Webpack for bundling.

## Prerequisites

- **Node.js** v16 or later (includes npm)
- **Outlook** that supports Office add-ins (Outlook on Windows, Mac, or Outlook on the web with an Office 365 account)
- A recent version of Office with [Mailbox 1.3](https://learn.microsoft.com/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) or later

## Installation

1. Clone the repository and navigate into it.
2. Run `npm install` to install all dependencies.
3. (Optional) Install development certificates by running `npx office-addin-dev-certs install` to avoid browser security prompts when developing locally.

## Development

- Start the dev server with:

  ```bash
  npm run dev-server
  ```

  This builds the project in development mode and serves it at <https://localhost:3000>.

- To build once in development mode, use:

  ```bash
  npm run build:dev
  ```

- To sideload the add-in in Outlook while developing, run:

  ```bash
  npm start
  ```

  Outlook will launch with the add-in automatically installed. Stop the session with `npm stop`.

## Production Build

1. Update the `urlProd` setting in `webpack.config.js` to point at the production host for your add-in files.
2. Build the production bundle:

   ```bash
   npm run build
   ```

   The output files are placed in the `dist` folder. Deploy these files to a web server.
3. Update the URLs in `manifest.xml` if necessary and distribute the manifest for installation.

## Sideloading the Add-in

For local development you can run `npm start` which sideloads the add-in automatically. To sideload manually:

1. Build the add-in (`npm run build:dev` or `npm run build`).
2. Open Outlook and choose **Get Add-ins** → **My Add-ins** → **Add a custom add-in** → **Add from file**.
3. Select the `manifest.xml` file in this project and confirm.

After sideloading, the add-in will appear in your Outlook ribbon under the **Synergia File Saver** group.

