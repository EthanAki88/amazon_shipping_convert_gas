/**
 * Event-based file monitoring for Amazon shipping data processing
 * Uses Google Drive API webhooks to detect file uploads
 */

const INPUT_FOLDER_ID = '1aIy5uWOuMkiFQM5aZPl1avns_odHx2r4';
const CHANNEL_ID = 'amazon-shipping-event-channel';

/**
 * Webhook endpoint for Google Drive API push notifications
 */
function doPost(e) {
  Logger.log('Received Drive notification: ' + JSON.stringify(e));

  const resourceState = e.headers['X-Goog-Resource-State'];
  const channelId = e.headers['X-Goog-Channel-Id'];

  Logger.log(`Resource State: ${resourceState}`);
  Logger.log(`Channel ID: ${channelId}`);

  // Verify channel ID
  const storedChannelId = PropertiesService.getUserProperties().getProperty('EVENT_CHANNEL_ID');
  if (channelId !== storedChannelId) {
    Logger.log('Mismatched channel ID. Ignoring notification.');
    return;
  }

  // Process change events
  if (resourceState === 'change') {
    Logger.log('File change detected! Checking for required files...');
    handleFileChange();
  } else if (resourceState === 'sync') {
    Logger.log('Drive watch channel synchronized.');
  } else if (resourceState === 'not_found' || resourceState === 'gone') {
    Logger.log('Drive watch channel expired. Re-establishing watch.');
    setupEventWatch();
  }
}

/**
 * Handle file changes and trigger processing if all required files are present
 */
function handleFileChange() {
  try {
    const fileStatus = checkRequiredFiles();
    
    if (fileStatus.allFilesPresent) {
      Logger.log('✅ All required files detected! Triggering processing...');
      
      // Check if files are recent (within last 5 minutes)
      const now = new Date();
      const fiveMinutesAgo = new Date(now.getTime() - 5 * 60 * 1000);
      
      let hasRecentFiles = false;
      for (const file of fileStatus.files) {
        if (file.lastModified && file.lastModified > fiveMinutesAgo) {
          hasRecentFiles = true;
          Logger.log(`Recent file detected: ${file.name}`);
          break;
        }
      }
      
      if (hasRecentFiles) {
        Logger.log('🚀 Recent file uploads detected! Starting processing...');
        try {
          processAmazonShippingData();
          Logger.log('✅ Processing completed successfully!');
        } catch (error) {
          Logger.log('❌ Error in processAmazonShippingData: ' + error.message);
        }
      } else {
        Logger.log('📁 Files are present but not recent. Waiting for new uploads...');
      }
    } else {
      Logger.log('⏳ Missing required files. Waiting for uploads...');
      Logger.log('Missing files: ' + fileStatus.missingFiles.join(', '));
    }
    
  } catch (error) {
    Logger.log('❌ Error handling file change: ' + error.message);
  }
}

/**
 * Check if all required files are present in the input folder
 */
function checkRequiredFiles() {
  try {
    const folder = DriveApp.getFolderById(INPUT_FOLDER_ID);
    const files = folder.getFiles();
    
    const requiredFiles = {
      '佐川.csv': { found: false, file: null, lastModified: null },
      '昭新紙業.csv': { found: false, file: null, lastModified: null },
      '福山通運.csv': { found: false, file: null, lastModified: null },
      'amazon_txt': { found: false, file: null, lastModified: null, pattern: /^\d{17}\.txt$/ }
    };
    
    const allFiles = [];
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const lastModified = file.getLastUpdated();
      
      allFiles.push({
        name: fileName,
        lastModified: lastModified,
        size: file.getSize()
      });
      
      // Check for required CSV files
      if (requiredFiles.hasOwnProperty(fileName)) {
        requiredFiles[fileName].found = true;
        requiredFiles[fileName].file = file;
        requiredFiles[fileName].lastModified = lastModified;
      }
      
      // Check for Amazon TXT file (17-digit filename)
      if (requiredFiles.amazon_txt.pattern.test(fileName)) {
        requiredFiles.amazon_txt.found = true;
        requiredFiles.amazon_txt.file = file;
        requiredFiles.amazon_txt.lastModified = lastModified;
      }
    }
    
    // Check if all required files are present
    let allFilesPresent = true;
    const missingFiles = [];
    
    for (const [fileName, status] of Object.entries(requiredFiles)) {
      if (fileName === 'amazon_txt') {
        if (!status.found) {
          allFilesPresent = false;
          missingFiles.push('Amazon TXT file (17-digit filename)');
        }
      } else {
        if (!status.found) {
          allFilesPresent = false;
          missingFiles.push(fileName);
        }
      }
    }
    
    return {
      allFilesPresent: allFilesPresent,
      files: allFiles,
      missingFiles: missingFiles,
      requiredFiles: requiredFiles
    };
    
  } catch (error) {
    Logger.log('❌ Error checking required files: ' + error.message);
    return {
      allFilesPresent: false,
      files: [],
      missingFiles: ['Error checking files'],
      requiredFiles: {}
    };
  }
}

/**
 * Set up event-based file monitoring using Drive API webhooks
 */
function setupEventWatch() {
  const userProperties = PropertiesService.getUserProperties();
  const scriptUrl = ScriptApp.getService().getUrl();
  
  if (!scriptUrl) {
    Logger.log('Cannot setup event watch - Web App URL not available');
    return;
  }

  // Check if Drive API is available
  try {
    const testResponse = Drive.Changes.getStartPageToken();
    Logger.log('✅ Drive API is available');
  } catch (error) {
    Logger.log('❌ Drive API not available: ' + error.message);
    Logger.log('Please enable the Drive API service in your Apps Script project');
    return;
  }

  // Define the watch request for Drive API v2
  const watchRequest = {
    id: CHANNEL_ID,
    type: 'web_hook', // Correct type for Drive API v2
    address: scriptUrl,
    expiration: (new Date().getTime() + 3600000).toString(), // 1 hour expiration
  };

  try {
    // Get the current startPageToken
    const startPageTokenResponse = Drive.Changes.getStartPageToken();
    const startPageToken = startPageTokenResponse.startPageToken;
    
    userProperties.setProperty('EVENT_START_PAGE_TOKEN', startPageToken);
    userProperties.setProperty('EVENT_CHANNEL_ID', CHANNEL_ID);
    Logger.log('Current startPageToken: ' + startPageToken);

    // Set up the webhook using Drive API v2
    const response = Drive.Changes.watch({
      pageToken: startPageToken,
      resource: watchRequest
    });

    userProperties.setProperty('EVENT_RESOURCE_ID', response.resourceId);
    Logger.log('✅ Successfully set up event-based file monitoring');
    Logger.log('Webhook response: ' + JSON.stringify(response));
    Logger.log('Notifications will be sent to: ' + scriptUrl);
    Logger.log(`Watching for file uploads in folder ID: ${INPUT_FOLDER_ID}`);

  } catch (error) {
    Logger.log('❌ Error setting up event watch: ' + error.message);
    Logger.log('Error details: ' + error.toString());
    
    if (error.message.includes('Unknown channel type')) {
      Logger.log('This error usually means:');
      Logger.log('1. Drive API service is not enabled');
      Logger.log('2. The script is not deployed as a Web App');
      Logger.log('3. The Web App URL is not accessible');
      Logger.log('4. The webhook type format is incorrect');
      Logger.log('');
      Logger.log('To fix this:');
      Logger.log('1. Go to Deploy > New deployment');
      Logger.log('2. Choose "Web app" as type');
      Logger.log('3. Set "Execute as" to your account');
      Logger.log('4. Set "Who has access" to "Anyone"');
      Logger.log('5. Deploy and use the new URL');
    }
  }
}

/**
 * Stop the event-based file monitoring
 */
function stopEventWatch() {
  const userProperties = PropertiesService.getUserProperties();
  const storedResourceId = userProperties.getProperty('EVENT_RESOURCE_ID');
  const storedChannelId = userProperties.getProperty('EVENT_CHANNEL_ID');

  if (!storedResourceId || !storedChannelId) {
    Logger.log('No resource ID or channel ID found to stop the watch');
    return;
  }

  try {
    Drive.Channels.stop({
      id: storedChannelId,
      resourceId: storedResourceId
    });
    Logger.log('✅ Event-based file monitoring stopped successfully');
    
    // Clear stored properties
    userProperties.deleteProperty('EVENT_RESOURCE_ID');
    userProperties.deleteProperty('EVENT_START_PAGE_TOKEN');
    userProperties.deleteProperty('EVENT_CHANNEL_ID');
  } catch (error) {
    Logger.log('❌ Error stopping event watch: ' + error.message);
  }
}

/**
 * Alternative setup function that tries different webhook formats
 */
function setupEventWatchAlternative() {
  const userProperties = PropertiesService.getUserProperties();
  const scriptUrl = ScriptApp.getService().getUrl();
  
  if (!scriptUrl) {
    Logger.log('Cannot setup event watch - Web App URL not available');
    return;
  }

  // Get the current startPageToken
  const startPageTokenResponse = Drive.Changes.getStartPageToken();
  const startPageToken = startPageTokenResponse.startPageToken;
  
  userProperties.setProperty('EVENT_START_PAGE_TOKEN', startPageToken);
  userProperties.setProperty('EVENT_CHANNEL_ID', CHANNEL_ID);
  Logger.log('Current startPageToken: ' + startPageToken);

  // Try different webhook formats
  const formats = [
    { type: 'web_hook', description: 'Drive API v2 format' },
    { type: 'webhook', description: 'Legacy format' },
    { type: 'http', description: 'HTTP format' }
  ];
  
  for (const format of formats) {
    try {
      Logger.log(`Trying ${format.description}...`);
      
      const watchRequest = {
        id: CHANNEL_ID + '_' + format.type,
        type: format.type,
        address: scriptUrl,
        expiration: (new Date().getTime() + 3600000).toString(),
      };
      
      const response = Drive.Changes.watch({
        pageToken: startPageToken,
        resource: watchRequest
      });
      
      Logger.log(`✅ Success with ${format.description}: ${JSON.stringify(response)}`);
      
      // Store the successful configuration
      userProperties.setProperty('EVENT_RESOURCE_ID', response.resourceId);
      userProperties.setProperty('EVENT_CHANNEL_ID', CHANNEL_ID + '_' + format.type);
      
      Logger.log('✅ Successfully set up event-based file monitoring');
      Logger.log('Notifications will be sent to: ' + scriptUrl);
      Logger.log(`Watching for file uploads in folder ID: ${INPUT_FOLDER_ID}`);
      
      return true;
      
    } catch (error) {
      Logger.log(`❌ Failed with ${format.description}: ${error.message}`);
    }
  }
  
  Logger.log('❌ All webhook formats failed. Please check your deployment settings.');
  return false;
}

/**
 * Test Drive API availability and webhook setup
 */
function testDriveAPI() {
  Logger.log('=== Testing Drive API and Webhook Setup ===');
  
  // Test Drive API availability
  try {
    const testResponse = Drive.Changes.getStartPageToken();
    Logger.log('✅ Drive API is available');
    Logger.log('Start page token: ' + testResponse.startPageToken);
  } catch (error) {
    Logger.log('❌ Drive API not available: ' + error.message);
    Logger.log('Please enable the Drive API service in your Apps Script project');
    return false;
  }
  
  // Test Web App URL
  const scriptUrl = ScriptApp.getService().getUrl();
  Logger.log('Web App URL: ' + scriptUrl);
  
  if (scriptUrl.includes('/dev')) {
    Logger.log('⚠️  Warning: Using development URL');
    Logger.log('Please deploy as production Web App for webhook functionality');
  } else {
    Logger.log('✅ Using production URL');
  }
  
  // Test webhook URL accessibility
  try {
    const response = UrlFetchApp.fetch(scriptUrl, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    Logger.log(`Webhook URL response code: ${responseCode}`);
    
    if (responseCode === 200) {
      Logger.log('✅ Webhook URL is accessible');
    } else {
      Logger.log('❌ Webhook URL returned error code: ' + responseCode);
    }
    
  } catch (error) {
    Logger.log('❌ Error testing webhook URL: ' + error.message);
  }
  
  return true;
}

/**
 * Complete setup function for event-based monitoring
 */
function completeEventSetup() {
  Logger.log('=== Event-Based File Monitoring Setup ===');
  
  // Test Drive API first
  if (!testDriveAPI()) {
    Logger.log('❌ Drive API test failed. Please fix issues above.');
    return;
  }
  
  // Check Web App deployment
  const scriptUrl = ScriptApp.getService().getUrl();
  if (scriptUrl.includes('/dev')) {
    Logger.log('❌ Using development URL. Please deploy as production Web App.');
    Logger.log('1. Click "Deploy" > "New deployment"');
    Logger.log('2. Choose "Web app" as type');
    Logger.log('3. Set "Execute as" to "Me"');
    Logger.log('4. Set "Who has access" to "Anyone"');
    Logger.log('5. Deploy and use the new URL');
    return;
  }
  
  Logger.log('✅ Using production URL: ' + scriptUrl);
  
  // Try standard setup first
  Logger.log('');
  Logger.log('Attempting standard webhook setup...');
  setupEventWatch();
  
  // If standard setup fails, try alternative
  const userProperties = PropertiesService.getUserProperties();
  const resourceId = userProperties.getProperty('EVENT_RESOURCE_ID');
  
  if (!resourceId) {
    Logger.log('');
    Logger.log('Standard setup failed. Trying alternative webhook formats...');
    setupEventWatchAlternative();
  }
  
  Logger.log('');
  Logger.log('=== Event-Based Setup Complete ===');
  Logger.log('Your event-based file monitoring is now active!');
  Logger.log('Upload files to the input folder to trigger automatic processing.');
}

/**
 * Test the current file status
 */
function testFileStatus() {
  Logger.log('=== Testing Current File Status ===');
  
  const fileStatus = checkRequiredFiles();
  
  Logger.log(`All required files present: ${fileStatus.allFilesPresent}`);
  Logger.log(`Total files in input folder: ${fileStatus.files.length}`);
  
  if (fileStatus.missingFiles.length > 0) {
    Logger.log('Missing files: ' + fileStatus.missingFiles.join(', '));
  }
  
  Logger.log('');
  Logger.log('All files in input folder:');
  fileStatus.files.forEach(file => {
    const dateStr = file.lastModified.toLocaleString();
    Logger.log(`- ${file.name} (${file.size} bytes, modified: ${dateStr})`);
  });
}
