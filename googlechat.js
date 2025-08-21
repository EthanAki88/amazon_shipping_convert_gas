
/**
 * Send message to Google Chat
 * @param {string} message - The message to send
 * @param {string} webhookUrl - Google Chat webhook URL (optional, uses default if not provided)
 * @param {string} threadKey - Thread key for replies (optional)
 */
function sendToGoogleChat(message, webhookUrl = null, threadKey = null) {
    try {
      // Default webhook URL - replace with your actual webhook URL
      const defaultWebhookUrl = 'YOUR_GOOGLE_CHAT_WEBHOOK_URL_HERE';
      const url = webhookUrl || defaultWebhookUrl;
      
      if (url === 'YOUR_GOOGLE_CHAT_WEBHOOK_URL_HERE') {
        console.error('Google Chat webhook URL not configured. Please set your webhook URL.');
        return false;
      }
      
      // Prepare the message payload
      const payload = {
        text: message
      };
      
      // Add thread key if provided (for replies)
      if (threadKey) {
        payload.thread = {
          threadKey: threadKey
        };
      }
      
      // Send the message
      const options = {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload)
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode === 200) {
        console.log('✅ Google Chatメッセージを送信しました');
        return true;
      } else {
        console.error(`❌ Google Chat送信エラー: HTTP ${responseCode}`);
        console.error(`レスポンス: ${response.getContentText()}`);
        return false;
      }
      
    } catch (error) {
      console.error('Google Chat送信中のエラー:', error.message);
      return false;
    }
  }
  
  /**
   * Send formatted notification to Google Chat
   * @param {string} title - Message title
   * @param {string} content - Message content
   * @param {string} status - Status (success, error, warning, info)
   * @param {string} webhookUrl - Webhook URL (optional)
   */
  function sendNotificationToChat(title, content, status = 'info', webhookUrl = null) {
    try {
      // Status emojis and colors
      const statusConfig = {
        success: { emoji: '✅', color: '#00FF00' },
        error: { emoji: '❌', color: '#FF0000' },
        warning: { emoji: '⚠️', color: '#FFA500' },
        info: { emoji: 'ℹ️', color: '#0000FF' }
      };
      
      const config = statusConfig[status] || statusConfig.info;
      
      // Format the message
      const timestamp = new Date().toLocaleString('ja-JP');
      const formattedMessage = `${config.emoji} **${title}**\n\n${content}\n\n_${timestamp}_`;
      
      // Send the message
      return sendToGoogleChat(formattedMessage, webhookUrl);
      
    } catch (error) {
      console.error('通知送信中のエラー:', error.message);
      return false;
    }
  }
  
  /**
   * Send processing status to Google Chat
   * @param {string} status - Processing status
   * @param {object} details - Additional details
   * @param {string} webhookUrl - Webhook URL (optional)
   */
  function sendProcessingStatus(status, details = {}, webhookUrl = null) {
    try {
      let title, content, statusType;
      
      switch (status) {
        case 'started':
          title = '処理開始';
          content = 'Amazon出荷データの処理を開始しました。';
          statusType = 'info';
          break;
          
        case 'completed':
          title = '処理完了';
          content = `✅ Amazon出荷データの処理が完了しました。\n\n処理件数: ${details.processedCount || 0}件\n出力ファイル: ${details.outputFile || 'N/A'}`;
          statusType = 'success';
          break;
          
        case 'error':
          title = '処理エラー';
          content = `❌ 処理中にエラーが発生しました。\n\nエラー: ${details.error || 'Unknown error'}`;
          statusType = 'error';
          break;
          
        case 'files_missing':
          title = 'ファイル不足';
          content = `⚠️ 必要なファイルが不足しています。\n\n不足ファイル: ${details.missingFiles || 'Unknown'}`;
          statusType = 'warning';
          break;
          
        default:
          title = '処理状況';
          content = details.message || '処理状況の更新があります。';
          statusType = 'info';
      }
      
      return sendNotificationToChat(title, content, statusType, webhookUrl);
      
    } catch (error) {
      console.error('処理状況送信中のエラー:', error.message);
      return false;
    }
  }
  
  /**
   * Test Google Chat integration
   * @param {string} webhookUrl - Webhook URL to test (optional)
   */
  function testGoogleChat(webhookUrl = null) {
    try {
      console.log('Google Chat接続をテスト中...');
      
      const testMessage = '🧪 これはGoogle Chat接続のテストメッセージです。\n\n時刻: ' + new Date().toLocaleString('ja-JP');
      
      const success = sendToGoogleChat(testMessage, webhookUrl);
      
      if (success) {
        console.log('✅ Google Chat接続テスト成功');
        return true;
      } else {
        console.log('❌ Google Chat接続テスト失敗');
        return false;
      }
      
    } catch (error) {
      console.error('Google Chatテスト中のエラー:', error.message);
      return false;
    }
  }
  
  