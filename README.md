# Parse Excel File - Azure Function

A TypeScript Azure Function application for reading and transforming Excel files from SharePoint using ExcelJS.

## Features

- **Excel File Parsing**: Parse Excel files (.xlsx) and transform data into structured JSON format
- **SharePoint Integration**: Ready-to-use endpoint for SharePoint file access (requires additional configuration)
- **Data Transformation**: Built-in utilities for filtering, mapping, and transforming Excel data
- **TypeScript Support**: Full TypeScript implementation with type definitions
- **Azure Functions v4**: Built on the latest Azure Functions programming model

## Prerequisites

- Node.js 18.x or higher
- Azure Functions Core Tools v4
- Azure subscription (for deployment)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/hjohnsondev/parse-excel-file.git
cd parse-excel-file
```

2. Install dependencies:
```bash
npm install
```

3. Configure local settings:
Update `local.settings.json` with your SharePoint credentials (if using SharePoint integration):
```json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    "FUNCTIONS_WORKER_RUNTIME": "node",
    "SHAREPOINT_SITE_URL": "https://your-tenant.sharepoint.com/sites/your-site",
    "SHAREPOINT_CLIENT_ID": "your-client-id",
    "SHAREPOINT_CLIENT_SECRET": "your-client-secret"
  }
}
```

## Building the Project

Build TypeScript files:
```bash
npm run build
```

Watch mode for development:
```bash
npm run watch
```

## Running Locally

Start the Azure Functions runtime:
```bash
npm start
```

The function will be available at:
- `http://localhost:7071/api/parseExcelFile` - Parse uploaded Excel files
- `http://localhost:7071/api/readFromSharePoint` - Read files from SharePoint

## API Endpoints

### 1. Parse Excel File (POST /api/parseExcelFile)

Upload and parse an Excel file directly.

**Request:**
- Method: `POST`
- Content-Type: `application/octet-stream`
- Body: Excel file binary data

**Response:**
```json
{
  "success": true,
  "worksheetCount": 2,
  "data": {
    "worksheets": [
      {
        "name": "Sheet1",
        "rowCount": 10,
        "columnCount": 5,
        "rows": [
          {
            "Column1": "Value1",
            "Column2": "Value2"
          }
        ]
      }
    ]
  }
}
```

**Example using cURL:**
```bash
curl -X POST http://localhost:7071/api/parseExcelFile \
  -H "Content-Type: application/octet-stream" \
  --data-binary @path/to/file.xlsx
```

### 2. Read from SharePoint (POST /api/readFromSharePoint)

Request to read an Excel file from SharePoint.

**Request:**
- Method: `POST`
- Content-Type: `application/json`
- Body:
```json
{
  "siteUrl": "https://your-tenant.sharepoint.com/sites/your-site",
  "fileName": "data.xlsx",
  "folderPath": "/Shared Documents",
  "clientId": "optional-override",
  "clientSecret": "optional-override"
}
```

**Response:**
```json
{
  "message": "SharePoint integration endpoint ready",
  "instructions": {
    "description": "To complete SharePoint integration, install @pnp/sp library",
    "steps": [
      "Run: npm install @pnp/sp @pnp/nodejs",
      "Implement authentication using client credentials",
      "Download file from SharePoint document library",
      "Pass file buffer to ExcelJS for parsing"
    ]
  }
}
```

## Usage Examples

### Using the ExcelTransformer Utility

```typescript
import { ExcelTransformer } from './utils/excelTransformer';
import * as ExcelJS from 'exceljs';

// Load workbook
const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile('data.xlsx');

// Transform with options
const data = ExcelTransformer.transformWorkbook(workbook, {
  includeEmptyRows: false,
  trimValues: true,
  startRow: 1
});

// Filter rows
const filtered = ExcelTransformer.filterRows(data, row => row.Status === 'Active');

// Get specific worksheet
const sheet = ExcelTransformer.getWorksheet(data, 'Sheet1');
```

## Project Structure

```
parse-excel-file/
├── src/
│   ├── functions/
│   │   ├── parseExcelFile.ts      # Main Excel parsing function
│   │   └── readFromSharePoint.ts  # SharePoint integration endpoint
│   └── utils/
│       ├── excelTransformer.ts    # Excel transformation utilities
│       └── types.ts                # TypeScript type definitions
├── dist/                           # Compiled JavaScript output
├── host.json                       # Azure Functions host configuration
├── local.settings.json             # Local development settings
├── package.json                    # Node.js dependencies
└── tsconfig.json                   # TypeScript configuration
```

## Development

### Adding New Transformations

To add custom data transformations:

1. Create a new transformation function in `src/utils/excelTransformer.ts`
2. Add necessary type definitions in `src/utils/types.ts`
3. Import and use in your function handlers

### Testing

Run tests (when implemented):
```bash
npm test
```

## Deployment

Deploy to Azure:

1. Create an Azure Function App
2. Configure application settings (same as local.settings.json)
3. Deploy using Azure Functions Core Tools:
```bash
func azure functionapp publish <your-function-app-name>
```

Or use VS Code Azure Functions extension for deployment.

## SharePoint Integration

To enable full SharePoint integration:

1. Install PnP libraries:
```bash
npm install @pnp/sp @pnp/nodejs
```

2. Register an app in Azure AD with SharePoint permissions
3. Configure client ID and secret in settings
4. Implement authentication and file download logic

## Security Considerations

- Never commit `local.settings.json` with real credentials
- Use Azure Key Vault for production secrets
- Implement proper authentication levels for production
- Validate and sanitize all input data

## License

ISC

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.

## Support

For issues and questions, please open an issue on GitHub.