# Excel WebSocket Office Add-in (Shared Runtime)

A modern, high-performance Microsoft Excel Add-in engineered with **TypeScript**, **SolidJS**, and the **Office.js Shared Runtime**. It exposes a bidirectional local WebSocket bridge (`ws://127.0.0.1:8890`) allowing external desktop automation controllers, keyboard managers, and command-line tools to drive Excel headlessly.

```text
+-----------------------------------------------------------------------------------+
|                        EXCEL WEBSOCKET ARCHITECTURAL TOPOLOGY                     |
+-----------------------------------------------------------------------------------+

   +--------------------------+               +--------------------------+
   |  KEYBOARD / AUTOMATION   |               |   LOCAL CLI / SCRIPTS    |
   +--------------------------+               +--------------------------+
                |                                          |
                +--------------------+---------------------+
                                     |
                                     | Local WebSocket (ws://127.0.0.1:8890)
                                     v
                  +--------------------------------------+
                  |         EXCEL SHARED RUNTIME         |
                  |           lifetime: "long"           |
                  +--------------------------------------+
                                     |
               +---------------------+---------------------+
               |                                           |
               v                                           v
+-----------------------------+             +-----------------------------+
|    HEADLESS / BACKGROUND    |             |     TASKPANE (OPTIONAL)     |
+-----------------------------+             +-----------------------------+
| * WebSocket Client Listener |             | * Connection Settings & Port|
| * Headless Excel Operations |             | * Live Event Diagnostic Log |
| * Spotlight & Position Slots|             | * Spotlight Palette Controls|
| * Ribbon Status Dynamic Sync|             | * Non-intrusive on demand   |
+-----------------------------+             +-----------------------------+
```

---

## Key Features

- **Shared Runtime Persistence (`lifetime: "long"`)**: The add-in initializes once and remains active in the background for the entire Excel session. Headless operations execute without requiring the taskpane to stay open.
- **Zero-Pollution Position Slots**: Save and restore cell coordinates instantly across sheets and workbooks stored cleanly inside `Office.context.document.settings`.
- **WPS-Style Spotlight Highlighting**: High-speed crosshair row/column focus powered by non-destructive conditional formatting rules with coordinate anchor matching and auto-detox on sheet transitions.
- **Border Color Cycling**: Rotate between vibrant border palettes with an automatic temporary visual shape badge.
- **Dynamic Ribbon Status**: Native Excel Ribbon button dynamically reports real-time connection status (Online / Offline).
- **Split-File Reactive i18n**: Full bilingual support (English `en-US` and Simplified Chinese `zh-CN`).

---

## Sideloading & Deployment

You can run this add-in in **two ways**:
1. **GitHub Pages (Zero Local Server)**: Host static files on GitHub Pages; Excel fetches code over HTTPS from GitHub.
2. **Local Development Server (`localhost:3000`)**: Run locally with live-reload for rapid feature development.

---

### Option 1: GitHub Pages Deployment (Zero Local Node Server)

This repository includes a preconfigured GitHub Actions workflow (`.github/workflows/deploy.yml`) that automatically builds and deploys to **GitHub Pages** on every push to `main`.

#### Step 1: Enable GitHub Pages
1. Push this project to your GitHub repository (e.g., `https://github.com/<username>/<repo>`).
2. In your GitHub repository, navigate to **Settings** > **Pages**.
3. Under **Build and deployment** > **Source**, select **GitHub Actions**.

#### Step 2: Automatic Deployment
Whenever you push changes to `main`, GitHub Actions automatically:
- Compiles the Webpack bundle to `dist/`.
- Injects your GitHub Pages URL into `dist/manifest.xml`.
- Publishes the bundle to your live GitHub Pages endpoint:
  ```text
  https://<username>.github.io/<repo>/
  ```

#### Step 3: Sideload in Microsoft Excel
1. Download your deployed `manifest.xml` directly from:
   ```text
   https://<username>.github.io/<repo>/manifest.xml
   ```
2. Open Microsoft Excel Desktop (Windows or Web).
3. Go to **Insert** > **Add-ins** > **My Add-ins**.
4. Click **Manage My Add-ins** (top-right dropdown) > **Upload My Add-in**.
5. Select the downloaded `manifest.xml`.
6. The **Excel WebSocket** button will appear on your **Home** tab ribbon!

---

### Option 2: Local Development Setup

To develop or debug locally with instant live reload:

#### Prerequisites
- [Node.js](https://nodejs.org/) (v20+ or v22)
- Microsoft Excel Desktop (Office 2016+, Office 365, or Office LTSC)

#### Installation & Startup
```powershell
# 1. Install dependencies
npm install

# 2. Start local HTTPS dev server and auto-sideload into Excel
npm start
```

#### Available Commands
| Command | Description |
| :--- | :--- |
| `npm start` | Boots `webpack-dev-server` (`https://localhost:3000`) and launches Excel Desktop |
| `npm run dev-server` | Starts only the local Webpack dev server without launching Excel |
| `npm run build` | Compiles optimized production bundle into `./dist/` |
| `npm run build:gh-pages` | Builds production bundle and generates production manifests |
| `npm run stop` | Closes running debugging sessions and unregisters sideloaded manifest |
| `npm run validate` | Validates `manifest.xml` schema against official Office.js rules |

---

## WebSocket Protocol Specification

The add-in communicates with local automation daemons via standard JSON-RPC over WebSocket:

- **Target URL**: `ws://127.0.0.1:8890` (Configurable in Taskpane UI)
- **Channel**: `excel-addin-websocket`
- **Client Type**: `office-addin`
- **Client Name**: `excel`

### Message Structure Example
```json
{
  "type": "command",
  "action": "spotlight_toggle",
  "params": {
    "mode": "crosshair",
    "color": "#FFF3CD"
  }
}
```

---

## License
MIT
