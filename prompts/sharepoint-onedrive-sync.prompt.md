---
mode: 'agent'
description: 'Automated workflow for mirroring files to SharePoint and OneDrive using Microsoft Graph API with portable, cloud-vendor-independent architecture'
tools: ['codebase', 'terminalCommand']
model: 'gpt-4o'
---

# SharePoint and OneDrive File Synchronization

This workflow implements an automated, portable file synchronization solution that mirrors application files (like `dashboard-status.json` and `carpuncle.log`) to SharePoint and OneDrive using Microsoft Graph API, without requiring local storage dependencies.

## Prerequisites

Before implementing this synchronization workflow, ensure:

- **Microsoft 365 Subscription**: Active subscription with SharePoint and OneDrive access
- **App Registration**: Azure AD application registered with appropriate permissions
- **API Permissions**: Microsoft Graph API permissions configured
  - `Files.ReadWrite.All` - For OneDrive access
  - `Sites.ReadWrite.All` - For SharePoint access
- **Development Environment**: Node.js/Python/C# runtime (choose based on your stack)
- **No Local Drive Dependencies**: Solution must work without C: drive access

## Architecture Principles

### Portability Requirements
- ✅ **Cloud-Native**: No dependencies on specific local drives (C:, D:, etc.)
- ✅ **Container-Ready**: Can run in Docker, Kubernetes, or serverless environments
- ✅ **Platform-Independent**: Works on Windows, Linux, and macOS
- ✅ **Vendor-Neutral**: Uses standard OAuth2/REST APIs to avoid cloud lock-in

### Key Design Patterns
1. **In-Memory Processing**: Process files in memory when possible
2. **Temporary Storage**: Use OS-agnostic temp directories (`os.tmpdir()`, `Path.GetTempPath()`)
3. **Stream-Based Upload**: Stream files directly from source to destination
4. **Configuration as Code**: Store settings in environment variables or cloud configuration services

## Workflow Steps

### Step 1: Azure AD App Registration and Authentication Setup

**Action**: Register application and configure authentication

**Process**:
1. **Register Application in Azure Portal**:
   - Navigate to Azure Portal → Azure Active Directory → App Registrations
   - Create new registration with name: `FileSync-Application`
   - Set supported account types (single/multi-tenant based on needs)
   - Configure redirect URI if using interactive auth

2. **Configure API Permissions**:
   ```
   Microsoft Graph Permissions:
   - Files.ReadWrite.All (Application or Delegated)
   - Sites.ReadWrite.All (Application or Delegated)
   ```
   - Request admin consent for application permissions

3. **Create Client Secret**:
   - Navigate to Certificates & Secrets
   - Create new client secret
   - **IMPORTANT**: Store securely in Azure Key Vault or environment variables
   - Never commit secrets to source control

4. **Gather Required Credentials**:
   - Tenant ID: `{tenant-id}`
   - Client ID: `{client-id}`
   - Client Secret: `{client-secret}`

### Step 2: Implement Authentication Flow

**Action**: Establish secure authentication with Microsoft Graph API

**Implementation Options**:

#### Option A: Client Credentials Flow (Daemon/Service Apps)
Best for: Automated background services, no user interaction

**Node.js Example**:
```javascript
const msal = require('@azure/msal-node');

const config = {
    auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
        clientSecret: process.env.AZURE_CLIENT_SECRET,
    }
};

const cca = new msal.ConfidentialClientApplication(config);

async function getAccessToken() {
    const authResult = await cca.acquireTokenByClientCredential({
        scopes: ['https://graph.microsoft.com/.default'],
    });
    return authResult.accessToken;
}
```

**Python Example**:
```python
from msal import ConfidentialClientApplication
import os

app = ConfidentialClientApplication(
    client_id=os.environ['AZURE_CLIENT_ID'],
    client_credential=os.environ['AZURE_CLIENT_SECRET'],
    authority=f"https://login.microsoftonline.com/{os.environ['AZURE_TENANT_ID']}"
)

def get_access_token():
    result = app.acquire_token_for_client(
        scopes=['https://graph.microsoft.com/.default']
    )
    return result['access_token']
```

**C# Example**:
```csharp
using Microsoft.Identity.Client;

var app = ConfidentialClientApplicationBuilder
    .Create(Environment.GetEnvironmentVariable("AZURE_CLIENT_ID"))
    .WithClientSecret(Environment.GetEnvironmentVariable("AZURE_CLIENT_SECRET"))
    .WithAuthority(new Uri($"https://login.microsoftonline.com/{Environment.GetEnvironmentVariable("AZURE_TENANT_ID")}"))
    .Build();

var result = await app.AcquireTokenForClient(
    new[] { "https://graph.microsoft.com/.default" }
).ExecuteAsync();

string accessToken = result.AccessToken;
```

#### Option B: Managed Identity (Azure-Hosted Services)
Best for: Azure Functions, Azure App Service, Azure VMs

**Node.js Example**:
```javascript
const { DefaultAzureCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const credential = new DefaultAzureCredential();
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
});

const client = Client.initWithMiddleware({ authProvider });
```

### Step 3: Discover SharePoint Site and OneDrive Locations

**Action**: Identify target locations for file synchronization

**Process**:

1. **Get SharePoint Site Information**:
```javascript
// Node.js - Get site by name
async function getSharePointSite(siteName) {
    const response = await fetch(
        `https://graph.microsoft.com/v1.0/sites?search=${siteName}`,
        {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        }
    );
    const data = await response.json();
    return data.value[0]; // Returns site object with id
}

// Alternative: Get site by full URL
async function getSharePointSiteByUrl(siteUrl) {
    // For site: https://contoso.sharepoint.com/sites/ProjectTeam
    const hostname = 'contoso.sharepoint.com';
    const sitePath = '/sites/ProjectTeam';
    
    const response = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`,
        {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        }
    );
    return await response.json();
}
```

2. **Get Document Library (Drive)**:
```javascript
async function getDocumentLibrary(siteId, libraryName = 'Documents') {
    const response = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
        {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        }
    );
    const data = await response.json();
    
    // Find specific library by name
    const library = data.value.find(drive => drive.name === libraryName);
    return library;
}
```

3. **Get User's OneDrive**:
```javascript
// For current user (delegated permissions)
async function getMyOneDrive() {
    const response = await fetch(
        'https://graph.microsoft.com/v1.0/me/drive',
        {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        }
    );
    return await response.json();
}

// For specific user (application permissions)
async function getUserOneDrive(userId) {
    const response = await fetch(
        `https://graph.microsoft.com/v1.0/users/${userId}/drive`,
        {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        }
    );
    return await response.json();
}
```

### Step 4: Implement File Upload to SharePoint

**Action**: Upload files to SharePoint document library

**Implementation**:

#### Small Files (< 4MB) - Simple Upload
```javascript
async function uploadToSharePoint(driveId, fileName, fileContent, targetFolder = '') {
    const uploadUrl = targetFolder 
        ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${targetFolder}/${fileName}:/content`
        : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/${fileName}/content`;
    
    const response = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/octet-stream'
        },
        body: fileContent
    });
    
    if (!response.ok) {
        throw new Error(`Upload failed: ${response.statusText}`);
    }
    
    return await response.json();
}

// Example: Upload JSON file
const dashboardStatus = JSON.stringify({
    timestamp: new Date().toISOString(),
    status: 'healthy',
    metrics: { /* ... */ }
});

await uploadToSharePoint(driveId, 'dashboard-status.json', dashboardStatus, 'logs');
```

#### Large Files (> 4MB) - Resumable Upload Session
```javascript
async function uploadLargeFileToSharePoint(driveId, fileName, fileStream, targetFolder = '') {
    // Step 1: Create upload session
    const sessionUrl = targetFolder
        ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${targetFolder}/${fileName}:/createUploadSession`
        : `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${fileName}:/createUploadSession`;
    
    const sessionResponse = await fetch(sessionUrl, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            item: {
                '@microsoft.graph.conflictBehavior': 'replace'
            }
        })
    });
    
    const session = await sessionResponse.json();
    const uploadUrl = session.uploadUrl;
    
    // Step 2: Upload in chunks (recommended: 5-10 MB per chunk)
    const chunkSize = 5 * 1024 * 1024; // 5 MB
    const fileSize = fileStream.length;
    let bytesUploaded = 0;
    
    while (bytesUploaded < fileSize) {
        const chunkEnd = Math.min(bytesUploaded + chunkSize, fileSize);
        const chunk = fileStream.slice(bytesUploaded, chunkEnd);
        
        const chunkResponse = await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
                'Content-Length': chunk.length,
                'Content-Range': `bytes ${bytesUploaded}-${chunkEnd - 1}/${fileSize}`
            },
            body: chunk
        });
        
        if (!chunkResponse.ok && chunkResponse.status !== 202) {
            throw new Error(`Chunk upload failed: ${chunkResponse.statusText}`);
        }
        
        bytesUploaded = chunkEnd;
    }
    
    return { success: true, uploadedBytes: bytesUploaded };
}
```

### Step 5: Implement File Upload to OneDrive

**Action**: Upload files to OneDrive

**Implementation**:

The OneDrive upload process is identical to SharePoint, as both use the same Microsoft Graph Drive API:

```javascript
async function uploadToOneDrive(fileName, fileContent, targetFolder = '') {
    // Get user's OneDrive
    const oneDrive = await getMyOneDrive();
    const driveId = oneDrive.id;
    
    // Use the same upload function as SharePoint
    return await uploadToSharePoint(driveId, fileName, fileContent, targetFolder);
}

// Example: Upload log file
const logContent = fs.readFileSync('/app/logs/carpuncle.log', 'utf8');
await uploadToOneDrive('carpuncle.log', logContent, 'application-logs');
```

### Step 6: Implement Portable, Automated Synchronization Service

**Action**: Create a service that automatically syncs files without local drive dependencies

**Complete Node.js Implementation**:

```javascript
const { DefaultAzureCredential } = require('@azure/identity');
const fetch = require('node-fetch');
const os = require('os');
const path = require('path');
const fs = require('fs').promises;

class CloudFileSyncService {
    constructor(config) {
        this.tenantId = config.tenantId || process.env.AZURE_TENANT_ID;
        this.clientId = config.clientId || process.env.AZURE_CLIENT_ID;
        this.clientSecret = config.clientSecret || process.env.AZURE_CLIENT_SECRET;
        this.sharePointSiteId = config.sharePointSiteId;
        this.sharePointDriveId = config.sharePointDriveId;
        this.oneDriveDriveId = config.oneDriveDriveId;
        this.accessToken = null;
    }
    
    async authenticate() {
        const msal = require('@azure/msal-node');
        const cca = new msal.ConfidentialClientApplication({
            auth: {
                clientId: this.clientId,
                authority: `https://login.microsoftonline.com/${this.tenantId}`,
                clientSecret: this.clientSecret,
            }
        });
        
        const result = await cca.acquireTokenByClientCredential({
            scopes: ['https://graph.microsoft.com/.default'],
        });
        
        this.accessToken = result.accessToken;
    }
    
    async uploadFile(driveId, fileName, fileContent, targetFolder = '') {
        if (!this.accessToken) {
            await this.authenticate();
        }
        
        const uploadUrl = targetFolder 
            ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${targetFolder}/${fileName}:/content`
            : `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${fileName}:/content`;
        
        const response = await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${this.accessToken}`,
                'Content-Type': 'application/octet-stream'
            },
            body: fileContent
        });
        
        if (!response.ok) {
            throw new Error(`Upload failed: ${response.statusText}`);
        }
        
        return await response.json();
    }
    
    async syncToSharePoint(fileName, fileContent, targetFolder = '') {
        console.log(`Syncing ${fileName} to SharePoint...`);
        try {
            const result = await this.uploadFile(
                this.sharePointDriveId,
                fileName,
                fileContent,
                targetFolder
            );
            console.log(`✓ Successfully synced ${fileName} to SharePoint`);
            return result;
        } catch (error) {
            console.error(`✗ Failed to sync ${fileName} to SharePoint:`, error.message);
            throw error;
        }
    }
    
    async syncToOneDrive(fileName, fileContent, targetFolder = '') {
        console.log(`Syncing ${fileName} to OneDrive...`);
        try {
            const result = await this.uploadFile(
                this.oneDriveDriveId,
                fileName,
                fileContent,
                targetFolder
            );
            console.log(`✓ Successfully synced ${fileName} to OneDrive`);
            return result;
        } catch (error) {
            console.error(`✗ Failed to sync ${fileName} to OneDrive:`, error.message);
            throw error;
        }
    }
    
    async syncFileToBothLocations(fileName, fileContent, targetFolder = '') {
        const results = await Promise.allSettled([
            this.syncToSharePoint(fileName, fileContent, targetFolder),
            this.syncToOneDrive(fileName, fileContent, targetFolder)
        ]);
        
        return {
            sharePoint: results[0],
            oneDrive: results[1]
        };
    }
}

// Usage Example
async function main() {
    const syncService = new CloudFileSyncService({
        sharePointDriveId: 'b!-RIj2DuyvEyV1T4NlOaMHk...',
        oneDriveDriveId: 'b!CbtYWMQVdUGKv...'
    });
    
    // Example 1: Sync JSON dashboard status
    const dashboardStatus = JSON.stringify({
        timestamp: new Date().toISOString(),
        status: 'operational',
        services: {
            api: 'healthy',
            database: 'healthy'
        }
    }, null, 2);
    
    await syncService.syncFileToBothLocations(
        'dashboard-status.json',
        dashboardStatus,
        'monitoring'
    );
    
    // Example 2: Sync log file (using OS-agnostic temp directory)
    const tempDir = os.tmpdir();
    const logPath = path.join(tempDir, 'carpuncle.log');
    
    // Generate or read log content
    const logContent = `[${new Date().toISOString()}] Application started\n` +
                      `[${new Date().toISOString()}] All services operational\n`;
    
    // Optionally write to temp for processing, then upload
    await fs.writeFile(logPath, logContent);
    const logData = await fs.readFile(logPath, 'utf8');
    
    await syncService.syncFileToBothLocations(
        'carpuncle.log',
        logData,
        'logs'
    );
    
    // Clean up temp file
    await fs.unlink(logPath);
}

main().catch(console.error);
```

**Python Implementation**:

```python
import os
import json
import tempfile
from datetime import datetime
from msal import ConfidentialClientApplication
import requests

class CloudFileSyncService:
    def __init__(self, config):
        self.tenant_id = config.get('tenant_id') or os.environ['AZURE_TENANT_ID']
        self.client_id = config.get('client_id') or os.environ['AZURE_CLIENT_ID']
        self.client_secret = config.get('client_secret') or os.environ['AZURE_CLIENT_SECRET']
        self.sharepoint_drive_id = config['sharepoint_drive_id']
        self.onedrive_drive_id = config['onedrive_drive_id']
        self.access_token = None
        
    def authenticate(self):
        app = ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}"
        )
        
        result = app.acquire_token_for_client(
            scopes=['https://graph.microsoft.com/.default']
        )
        
        self.access_token = result['access_token']
    
    def upload_file(self, drive_id, file_name, file_content, target_folder=''):
        if not self.access_token:
            self.authenticate()
        
        if target_folder:
            upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{target_folder}/{file_name}:/content"
        else:
            upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_name}:/content"
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/octet-stream'
        }
        
        # Convert string content to bytes if necessary
        if isinstance(file_content, str):
            file_content = file_content.encode('utf-8')
        
        response = requests.put(upload_url, headers=headers, data=file_content)
        response.raise_for_status()
        
        return response.json()
    
    def sync_to_sharepoint(self, file_name, file_content, target_folder=''):
        print(f"Syncing {file_name} to SharePoint...")
        try:
            result = self.upload_file(
                self.sharepoint_drive_id,
                file_name,
                file_content,
                target_folder
            )
            print(f"✓ Successfully synced {file_name} to SharePoint")
            return result
        except Exception as e:
            print(f"✗ Failed to sync {file_name} to SharePoint: {str(e)}")
            raise
    
    def sync_to_onedrive(self, file_name, file_content, target_folder=''):
        print(f"Syncing {file_name} to OneDrive...")
        try:
            result = self.upload_file(
                self.onedrive_drive_id,
                file_name,
                file_content,
                target_folder
            )
            print(f"✓ Successfully synced {file_name} to OneDrive")
            return result
        except Exception as e:
            print(f"✗ Failed to sync {file_name} to OneDrive: {str(e)}")
            raise
    
    def sync_file_to_both_locations(self, file_name, file_content, target_folder=''):
        results = {}
        
        try:
            results['sharepoint'] = self.sync_to_sharepoint(file_name, file_content, target_folder)
        except Exception as e:
            results['sharepoint'] = {'error': str(e)}
        
        try:
            results['onedrive'] = self.sync_to_onedrive(file_name, file_content, target_folder)
        except Exception as e:
            results['onedrive'] = {'error': str(e)}
        
        return results

# Usage Example
def main():
    sync_service = CloudFileSyncService({
        'sharepoint_drive_id': 'b!-RIj2DuyvEyV1T4NlOaMHk...',
        'onedrive_drive_id': 'b!CbtYWMQVdUGKv...'
    })
    
    # Example 1: Sync JSON dashboard status
    dashboard_status = json.dumps({
        'timestamp': datetime.utcnow().isoformat(),
        'status': 'operational',
        'services': {
            'api': 'healthy',
            'database': 'healthy'
        }
    }, indent=2)
    
    sync_service.sync_file_to_both_locations(
        'dashboard-status.json',
        dashboard_status,
        'monitoring'
    )
    
    # Example 2: Sync log file (using OS-agnostic temp directory)
    with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.log') as temp_log:
        log_content = f"[{datetime.utcnow().isoformat()}] Application started\n"
        log_content += f"[{datetime.utcnow().isoformat()}] All services operational\n"
        temp_log.write(log_content)
        temp_log_path = temp_log.name
    
    # Read and upload
    with open(temp_log_path, 'r') as f:
        log_data = f.read()
    
    sync_service.sync_file_to_both_locations(
        'carpuncle.log',
        log_data,
        'logs'
    )
    
    # Clean up
    os.unlink(temp_log_path)

if __name__ == '__main__':
    main()
```

### Step 7: Implement Retry Logic and Error Handling

**Action**: Add resilience to handle transient failures

**Implementation**:

```javascript
class RetryableCloudFileSyncService extends CloudFileSyncService {
    constructor(config) {
        super(config);
        this.maxRetries = config.maxRetries || 3;
        this.retryDelay = config.retryDelay || 1000; // ms
    }
    
    async uploadFileWithRetry(driveId, fileName, fileContent, targetFolder = '', attempt = 1) {
        try {
            return await this.uploadFile(driveId, fileName, fileContent, targetFolder);
        } catch (error) {
            if (attempt < this.maxRetries) {
                console.log(`Retry attempt ${attempt}/${this.maxRetries} for ${fileName}...`);
                
                // Exponential backoff
                const delay = this.retryDelay * Math.pow(2, attempt - 1);
                await new Promise(resolve => setTimeout(resolve, delay));
                
                // Re-authenticate if token expired
                if (error.message.includes('401') || error.message.includes('403')) {
                    await this.authenticate();
                }
                
                return await this.uploadFileWithRetry(
                    driveId,
                    fileName,
                    fileContent,
                    targetFolder,
                    attempt + 1
                );
            }
            throw error;
        }
    }
    
    async syncToSharePoint(fileName, fileContent, targetFolder = '') {
        console.log(`Syncing ${fileName} to SharePoint...`);
        try {
            const result = await this.uploadFileWithRetry(
                this.sharePointDriveId,
                fileName,
                fileContent,
                targetFolder
            );
            console.log(`✓ Successfully synced ${fileName} to SharePoint`);
            return result;
        } catch (error) {
            console.error(`✗ Failed to sync ${fileName} to SharePoint after ${this.maxRetries} attempts:`, error.message);
            throw error;
        }
    }
}
```

### Step 8: Implement Scheduled Synchronization

**Action**: Set up automated, periodic synchronization

**Option A: Using Node.js Cron**:
```javascript
const cron = require('node-cron');

// Sync every hour
cron.schedule('0 * * * *', async () => {
    console.log('Starting scheduled file sync...');
    const syncService = new RetryableCloudFileSyncService({
        sharePointDriveId: process.env.SHAREPOINT_DRIVE_ID,
        oneDriveDriveId: process.env.ONEDRIVE_DRIVE_ID
    });
    
    // Generate current dashboard status
    const status = await generateDashboardStatus();
    await syncService.syncFileToBothLocations(
        'dashboard-status.json',
        JSON.stringify(status, null, 2),
        'monitoring'
    );
    
    console.log('Scheduled sync completed');
});
```

**Option B: Using Azure Functions (Timer Trigger)**:
```javascript
module.exports = async function (context, myTimer) {
    const syncService = new RetryableCloudFileSyncService({
        sharePointDriveId: process.env.SHAREPOINT_DRIVE_ID,
        oneDriveDriveId: process.env.ONEDRIVE_DRIVE_ID
    });
    
    const status = {
        timestamp: new Date().toISOString(),
        functionExecutionId: context.executionContext.invocationId,
        status: 'operational'
    };
    
    await syncService.syncFileToBothLocations(
        'dashboard-status.json',
        JSON.stringify(status, null, 2),
        'monitoring'
    );
    
    context.log('File sync completed successfully');
};
```

**function.json** (Azure Functions):
```json
{
  "bindings": [
    {
      "name": "myTimer",
      "type": "timerTrigger",
      "direction": "in",
      "schedule": "0 0 * * * *"
    }
  ]
}
```

### Step 9: Configure Environment Variables

**Action**: Set up portable configuration using environment variables

**Example `.env` file** (for local development):
```bash
# Azure AD Authentication
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret

# SharePoint Configuration
SHAREPOINT_SITE_NAME=YourSiteName
SHAREPOINT_DRIVE_ID=b!-RIj2DuyvEyV1T4NlOaMHk8XzniRzTdOwaDqC-VQUTp_KZdq0vqZT7jPGLOLvMVP

# OneDrive Configuration
ONEDRIVE_USER_ID=user@contoso.com
ONEDRIVE_DRIVE_ID=b!CbtYWMQVdUGKv4EtMsIvQ-m--CXJAHhNpqrI5K6Sp_g7e4E8AnQMSL8i_9aZzY9m

# Sync Configuration
SYNC_INTERVAL_MINUTES=60
SYNC_RETRY_COUNT=3
SYNC_TARGET_FOLDER=application-sync
```

**Docker Environment** (`docker-compose.yml`):
```yaml
version: '3.8'
services:
  file-sync:
    build: .
    environment:
      - AZURE_TENANT_ID=${AZURE_TENANT_ID}
      - AZURE_CLIENT_ID=${AZURE_CLIENT_ID}
      - AZURE_CLIENT_SECRET=${AZURE_CLIENT_SECRET}
      - SHAREPOINT_DRIVE_ID=${SHAREPOINT_DRIVE_ID}
      - ONEDRIVE_DRIVE_ID=${ONEDRIVE_DRIVE_ID}
    volumes:
      # No C: drive mounts - fully portable
      - /tmp/app-temp:/tmp
```

**Kubernetes Secret**:
```yaml
apiVersion: v1
kind: Secret
metadata:
  name: file-sync-secrets
type: Opaque
stringData:
  AZURE_TENANT_ID: "your-tenant-id"
  AZURE_CLIENT_ID: "your-client-id"
  AZURE_CLIENT_SECRET: "your-client-secret"
  SHAREPOINT_DRIVE_ID: "drive-id"
  ONEDRIVE_DRIVE_ID: "drive-id"
```

### Step 10: Implement Monitoring and Logging

**Action**: Add comprehensive monitoring and audit logging

**Implementation**:

```javascript
class MonitoredCloudFileSyncService extends RetryableCloudFileSyncService {
    constructor(config) {
        super(config);
        this.syncHistory = [];
    }
    
    async syncFileToBothLocations(fileName, fileContent, targetFolder = '') {
        const syncEvent = {
            fileName,
            targetFolder,
            timestamp: new Date().toISOString(),
            fileSize: Buffer.byteLength(fileContent),
            results: {}
        };
        
        try {
            const results = await super.syncFileToBothLocations(fileName, fileContent, targetFolder);
            syncEvent.results = results;
            syncEvent.status = 'completed';
            
            // Log successful sync
            this.logSyncEvent(syncEvent);
            
            // Optionally send to monitoring service
            await this.sendToMonitoring(syncEvent);
            
            return results;
        } catch (error) {
            syncEvent.status = 'failed';
            syncEvent.error = error.message;
            this.logSyncEvent(syncEvent);
            throw error;
        }
    }
    
    logSyncEvent(event) {
        this.syncHistory.push(event);
        
        // Console logging with structured format
        console.log(JSON.stringify({
            level: event.status === 'failed' ? 'ERROR' : 'INFO',
            message: `File sync ${event.status}`,
            ...event
        }));
    }
    
    async sendToMonitoring(event) {
        // Example: Send to Application Insights
        // const appInsights = require('applicationinsights');
        // appInsights.defaultClient.trackEvent({
        //     name: 'FileSyncCompleted',
        //     properties: event
        // });
        
        // Example: Send to custom monitoring endpoint
        // await fetch('https://monitoring.contoso.com/api/events', {
        //     method: 'POST',
        //     headers: { 'Content-Type': 'application/json' },
        //     body: JSON.stringify(event)
        // });
    }
    
    getSyncStatistics() {
        const total = this.syncHistory.length;
        const successful = this.syncHistory.filter(e => e.status === 'completed').length;
        const failed = this.syncHistory.filter(e => e.status === 'failed').length;
        
        return {
            totalSyncs: total,
            successful,
            failed,
            successRate: total > 0 ? (successful / total * 100).toFixed(2) + '%' : '0%',
            lastSync: this.syncHistory[this.syncHistory.length - 1]
        };
    }
}
```

## Avoiding Cloud Lock-In

### Abstraction Layer Pattern

Create an abstraction layer that allows swapping cloud providers:

```javascript
// Storage provider interface
class IStorageProvider {
    async authenticate() { throw new Error('Not implemented'); }
    async uploadFile(location, fileName, content) { throw new Error('Not implemented'); }
    async downloadFile(location, fileName) { throw new Error('Not implemented'); }
}

// Microsoft Graph implementation
class MicrosoftGraphStorageProvider extends IStorageProvider {
    async authenticate() { /* Implementation */ }
    async uploadFile(location, fileName, content) { /* Implementation */ }
}

// AWS S3 implementation (future-proofing)
class AWSS3StorageProvider extends IStorageProvider {
    async authenticate() { /* Implementation */ }
    async uploadFile(bucket, fileName, content) { /* Implementation */ }
}

// Factory pattern for provider selection
class StorageProviderFactory {
    static create(type) {
        switch (type) {
            case 'microsoft':
                return new MicrosoftGraphStorageProvider();
            case 'aws':
                return new AWSS3StorageProvider();
            default:
                throw new Error(`Unknown provider: ${type}`);
        }
    }
}

// Usage
const provider = StorageProviderFactory.create(process.env.STORAGE_PROVIDER || 'microsoft');
await provider.uploadFile('location', 'file.json', content);
```

## Best Practices Summary

### Security
- ✅ Store credentials in Azure Key Vault or secure environment variables
- ✅ Use Managed Identity when running in Azure
- ✅ Implement least-privilege access (only required Graph API permissions)
- ✅ Rotate client secrets regularly
- ✅ Enable audit logging for all file operations

### Portability
- ✅ No hardcoded drive paths (C:, D:, etc.)
- ✅ Use OS-agnostic temporary directories (`os.tmpdir()`, `Path.GetTempPath()`)
- ✅ Environment-based configuration
- ✅ Container-ready architecture
- ✅ Platform-independent code

### Reliability
- ✅ Implement retry logic with exponential backoff
- ✅ Handle token expiration gracefully
- ✅ Use resumable uploads for large files
- ✅ Implement health checks and monitoring
- ✅ Log all sync operations for audit trail

### Performance
- ✅ Use streaming for large files
- ✅ Implement batch operations when syncing multiple files
- ✅ Cache access tokens (valid for ~1 hour)
- ✅ Use parallel uploads to SharePoint and OneDrive when appropriate
- ✅ Optimize file sizes before upload (compression when applicable)

## Troubleshooting

### Common Issues

**Issue**: "401 Unauthorized" or "403 Forbidden"
- **Solution**: Verify API permissions are granted and admin consent is provided
- **Solution**: Check if access token has expired and re-authenticate

**Issue**: "404 Not Found" when uploading to SharePoint
- **Solution**: Verify drive ID and site ID are correct
- **Solution**: Ensure target folder exists or create it first

**Issue**: Upload fails for large files
- **Solution**: Use resumable upload session for files > 4MB
- **Solution**: Implement chunked upload with proper Content-Range headers

**Issue**: Rate limiting (429 Too Many Requests)
- **Solution**: Implement exponential backoff retry logic
- **Solution**: Reduce sync frequency
- **Solution**: Use batch operations to reduce API calls

## Success Criteria

- ✅ Files successfully sync to both SharePoint and OneDrive
- ✅ Solution runs without local drive dependencies
- ✅ Authentication works reliably with token refresh
- ✅ Retry logic handles transient failures
- ✅ Monitoring and logging capture all sync operations
- ✅ Solution is portable across Windows, Linux, and containers
- ✅ Abstraction layer allows future migration to other cloud providers
- ✅ All secrets stored securely (not in source code)

## Additional Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Upload files to SharePoint](https://docs.microsoft.com/en-us/graph/api/driveitem-put-content)
- [Upload large files with upload session](https://docs.microsoft.com/en-us/graph/api/driveitem-createuploadsession)
- [Microsoft Authentication Library (MSAL)](https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-overview)
- [Azure Managed Identity](https://docs.microsoft.com/en-us/azure/active-directory/managed-identities-azure-resources/)
- [Best practices for Microsoft Graph](https://docs.microsoft.com/en-us/graph/best-practices-concept)
