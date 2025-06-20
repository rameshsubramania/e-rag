# Teams SharePoint URL Builder

A Microsoft Teams tab app that automatically generates SharePoint site URLs based on Teams and channel names.

## Features

- Displays current Team and Channel information
- Automatically generates SharePoint site URLs following the standard format
- Copy URL functionality
- Works in any Teams channel (including private channels)

## GitHub Pages Setup

1. Fork this repository
2. Go to repository Settings > Pages
3. Under "Source", select "Deploy from a branch"
4. Select "gh-pages" branch and "/root" folder
5. Click Save
6. Wait a few minutes for the first deployment

The site will be available at: `https://YOUR_GITHUB_USERNAME.github.io/teams-sharepoint-url-builder/`

## Teams App Setup

1. Update the manifest file:
   - Open `/manifest/manifest.json`
   - Replace `{{REPLACE_WITH_NEW_GUID}}` with a new GUID
   - Replace `YOUR_GITHUB_USERNAME` with your actual GitHub username in:
     - `configurationUrl`
     - `validDomains`

2. Create the app package:
   - Copy the `/manifest` folder contents
   - Create a ZIP file containing:
     - `manifest.json`
     - `color.png`
     - `outline.png`

3. Upload to Teams:
   - Go to Teams Admin Portal > Teams apps > Manage apps
   - Click "Upload new app"
   - Select your ZIP file
   - Approve the app for your organization

## Using the App

1. Go to any Teams channel
2. Click the + button to add a tab
3. Select "SharePoint URL Builder"
4. The tab will show:
   - Team and Channel information
   - Generated SharePoint URL
   - Copy URL button

The SharePoint URL will follow the format:
`https://axleinfo.sharepoint.com/sites/[team-name]-[channel-name]`

## Development

This is a static HTML/JS application with no build requirements. All files in the `/docs` folder are served directly via GitHub Pages.
