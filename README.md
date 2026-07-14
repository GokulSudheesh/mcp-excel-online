# MCP Excel Online

An [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server that integrates with your OneDrive and Excel Workbooks, to enable creating and modifying work sheets using the [Microsoft Graph API](https://github.com/microsoftgraph/msgraph-sdk-python).

---

## 🚀 Quick Start (Using `uvx`)

Essentially the server runs in one line: `uvx mcp-excel-online@latest`.

This command will automatically download the latest code and run it. **It is recommended to always use `@latest`** to ensure you have the newest version with the latest features and bug fixes.

1. **☁️ Prerequisite: Azure App Registration**
   - You **must** configure an Azure Entra ID (Azure AD) app registration and grant the necessary Microsoft Graph API permissions first.
   - ➡️ Jump to the [**Azure App Registration**](#azure-app-registration) guide below.

2. **Install `uv`**
   - [Installation guide](https://docs.astral.sh/uv/getting-started/installation)

3. **🔑 Set Essential Environment Variables**
   - You need to tell the server how to authenticate. Set these variables in your terminal:
   - **(Linux/macOS)**

     ```bash
     export APP_CLIENT_ID="your-client-id"
     export APP_TENANT_ID="your-tenant-id"       # required for application permissions
     export APP_CLIENT_SECRET="your-client-secret" # required for application permissions
     export DRIVE_ID="your-drive-id"
     ```

   - **(Windows CMD)**

     ```cmd
     set APP_CLIENT_ID="your-client-id"
     set APP_TENANT_ID="your-tenant-id"
     set APP_CLIENT_SECRET="your-client-secret"
     set DRIVE_ID="your-drive-id"
     ```

   - **(Windows PowerShell)**

     ```powershell
     $env:APP_CLIENT_ID = "your-client-id"
     $env:APP_TENANT_ID = "your-tenant-id"
     $env:APP_CLIENT_SECRET = "your-client-secret"
     $env:DRIVE_ID = "your-drive-id"
     ```

   - ➡️ See [**Environment Variables**](#environment-variables) for the full list of options.

4. **🏃 Run the Server!**
   - `uvx` will automatically download and run the latest version of `mcp-excel-online`:

     ```bash
     uvx mcp-excel-online@latest
     ```

   - The server will start and print logs indicating it's ready.

   > **💡 Pro Tip:** Always use `@latest` to ensure you get the newest version with bug fixes and features. Without `@latest`, `uvx` may use a cached older version.

5. **🔌 Connect your MCP Client**
   - Configure your client (e.g., Claude Desktop) to connect to the running server.
   - Depending on the client you use, you might not need step 4 because the client can launch the server for you. But it's a good practice to test run step 4 anyway to make sure things are set up properly.

6. **⚡ Optional: Enable Tool Filtering (Reduce Context Usage)**
   - By default, all tools are enabled. To reduce context usage, enable only the tools you need.
   - ➡️ See [**Available tool names for `--include-tools`**](#available-tool-names-for---include-tools) for details.

Set up is complete, you can now start issuing commands via your MCP client.

---

## 📖 Table of Contents

- [MCP Excel Online](#mcp-excel-online)
  - [🚀 Quick Start (Using `uvx`)](#-quick-start-using-uvx)
  - [📖 Table of Contents](#-table-of-contents)
  - [Features](#features)
    - [Resources](#resources)
    - [Tools](#tools)
  - [Prerequisites](#prerequisites)
  - [Azure App Registration](#azure-app-registration)
    - [Permission Types](#permission-types)
  - [Environment Variables](#environment-variables)
  - [Arguments](#arguments)
    - [Available tool names for `--include-tools`](#available-tool-names-for---include-tools)
    - [Running Locally with uv](#running-locally-with-uv)
      - [1. Clone the repository](#1-clone-the-repository)
      - [2. Create your `.env` file with the required environment variables (see above)](#2-create-your-env-file-with-the-required-environment-variables-see-above)
      - [3. Start the server](#3-start-the-server)
    - [Running with Docker](#running-with-docker)
      - [1. Building the image locally](#1-building-the-image-locally)
      - [OR Pulling the image from Docker Hub](#or-pulling-the-image-from-docker-hub)
      - [2. Run the container](#2-run-the-container)

---

## Features

### Resources

| Resource                                                  | Description                                           |
| --------------------------------------------------------- | ----------------------------------------------------- |
| `drive://root/children`                                   | List all items at the root of the configured OneDrive |
| `drive://drives/{drive_id}/root/{relative_path}/children` | List items in a folder by path                        |

### Tools

| Tool                       | Description                                               |
| -------------------------- | --------------------------------------------------------- |
| **List Sheets**            | List all worksheets in a workbook                         |
| **Create Worksheet**       | Add a new worksheet to a workbook                         |
| **Rename Worksheet**       | Rename an existing worksheet                              |
| **Delete Worksheet**       | Delete a worksheet from a workbook                        |
| **Copy Worksheet**         | Copy a worksheet from one workbook to another             |
| **Get Worksheet Data**     | Read cell values from a worksheet (full sheet or a range) |
| **Get Worksheet Formulas** | Read formulas from a worksheet (full sheet or a range)    |
| **Update Worksheet Data**  | Write values to a cell range in a worksheet               |

---

## Prerequisites

- Python 3.13+
- [`uv`](https://github.com/astral-sh/uv) (for local development)
- Docker (for containerised deployment)
- A Microsoft Azure account with an Entra ID (Azure AD) app registration

---

## Azure App Registration

Register your application by following the steps at [Register your app with the Microsoft Identity Platform.](https://docs.microsoft.com/graph/auth-register-app-v2)

### Permission Types

**Application permissions** are used when you don't need a user to log in to your app. The app performs tasks on its own and runs in the background using a client secret (service principal flow).
Required: `APP_CLIENT_ID`, `APP_TENANT_ID`, `APP_CLIENT_SECRET`

**Delegated permissions** (scopes) are used when your app requires a user to log in and interact with data related to that user in a session. Authentication is performed via the device code flow — the server will print a URL and a code to the console.
Required: `APP_CLIENT_ID` only

---

## Environment Variables

| Variable            | Required                     | Description                                                |
| ------------------- | ---------------------------- | ---------------------------------------------------------- |
| `APP_CLIENT_ID`     | **Always**                   | The Application (client) ID of your Azure app registration |
| `APP_TENANT_ID`     | Application permissions only | The Directory (tenant) ID of your Azure tenant             |
| `APP_CLIENT_SECRET` | Application permissions only | A client secret generated for your app registration        |
| `DRIVE_ID`          | Required                     | OneDrive Drive ID to use as the default drive              |
| `HOST`              | Optional                     | Host address the server binds to (default: `0.0.0.0`)      |
| `PORT`              | Optional                     | Port the server listens on (default: `8000`)               |

Create a `.env` file in the project root for local development:

```env
APP_CLIENT_ID=your-client-id
APP_TENANT_ID=your-tenant-id        # required for application permissions
APP_CLIENT_SECRET=your-client-secret # required for application permissions
DRIVE_ID=your-drive-id
```

---

## Arguments

| Argument             | Options                           | Default     | Description                                                                      |
| -------------------- | --------------------------------- | ----------- | -------------------------------------------------------------------------------- |
| `--transport`        | `stdio`, `sse`, `streamable-http` | `sse`       | MCP transport protocol to use                                                    |
| `--graph-permission` | `delegated`, `application`        | `delegated` | Microsoft Graph permission type (see above)                                      |
| `--include-tools`    | Comma-separated tool names        | _(all)_     | Only register the specified tools. If omitted, all tools are enabled (see below) |

### Available tool names for `--include-tools`

| Tool name                | Description                                               |
| ------------------------ | --------------------------------------------------------- |
| `list_sheets`            | List all worksheets in a workbook                         |
| `create_sheet`           | Add a new worksheet to a workbook                         |
| `rename_sheet`           | Rename an existing worksheet                              |
| `delete_sheet`           | Delete a worksheet from a workbook                        |
| `copy_sheet`             | Copy a worksheet from one workbook to another             |
| `get_worksheet_data`     | Read cell values from a worksheet (full sheet or a range) |
| `get_worksheet_formulas` | Read formulas from a worksheet (full sheet or a range)    |
| `update_worksheet_data`  | Write values to a cell range in a worksheet               |

**Example** — expose only read-only tools:

```bash
uv run mcp-excel-online --include-tools list_sheets,get_worksheet_data,get_worksheet_formulas
```

---

### Running Locally with uv

#### 1. Clone the repository

```bash
git clone https://github.com/GokulSudheesh/mcp-excel-online.git
cd mcp-excel-online
```

#### 2. Create your `.env` file with the required environment variables (see above)

#### 3. Start the server

```bash
# SSE transport with delegated permissions (device code login)
uv run mcp-excel-online --transport sse --graph-permission delegated

# stdio transport with application permissions (service principal)
uv run mcp-excel-online --transport stdio --graph-permission application

# streamable-http transport
uv run mcp-excel-online --transport streamable-http

# Only enable specific tools
uv run mcp-excel-online --include-tools list_sheets,get_worksheet_data,get_worksheet_formulas
```

---

### Running with Docker

#### 1. Building the image locally

```bash
docker build -t mcp-excel-online .
```

#### OR Pulling the image from Docker Hub

```bash
docker pull batmansroomie/mcp-excel-online:main
```

#### 2. Run the container

```bash
docker run --rm -p 8000:8000 \
  -e APP_CLIENT_ID=your-client-id \
  -e DRIVE_ID=your-drive-id \
  mcp-excel-online
```

You can also pass a `.env` file:

```bash
docker run --rm -p 8000:8000 --env-file .env mcp-excel-online
```

The server is accessible at `http://localhost:8000/sse` (SSE transport).

---
