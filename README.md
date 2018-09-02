# OneNote Export

Export OneNote notebooks as JSON + HTML using the [Microsoft Graph](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/onenote-api-overview) REST API. The [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) is helpful for learning the API.

## Access Tokens

Most of the battle is getting access tokens for your personal Microsoft account. Try this "streamlined" process:

1. [Register an application](https://docs.microsoft.com/en-us/onedrive/developer/rest-api/getting-started/app-registration) with MS Graph
2. Under the **Platforms** header, create a web app and set the **Redirect URL** to `http://localhost:3000/token`
3. Set the environment variable `CLIENT_ID` to the **Application ID** generated for your app
4. Start the Node web app and visit `/login` in your browser
3. Login to your MS account, getting redirected to `/token`
5. Copy the access token out of the URL and assign it to the environment variable `ACCESS_TOKEN`
7. Restart the Node web app

Be warned that the access token expires after an hour or so.

## Exporting

The web app presents a simplified version of Microsoft's API for getting OneNote data:

- `/notebooks`: Get all notebooks as JSON list.
- `/sections`: Get all notebook sections as JSON list.
- `/pages`: Get all notebook pages (excluding HTML content) as JSON list.
- `/content`: Get all notebook pages (including HTML content) as JSON list.