const archiveFolderId = '1aIy5uWOuMkiFQM5aZPl1avns_odHx2r4';
const flatFileFolderId = '1aIy5uWOuMkiFQM5aZPl1avns_odHx2r4';
const outputFolderId = '1aIy5uWOuMkiFQM5aZPl1avns_odHx2r4';
const inputFolderId = '1aIy5uWOuMkiFQM5aZPl1avns_odHx2r4';

/**
 * Main function to process Amazon shipping data and create output file
 */
function processAmazonShippingData() {
  // let filesToArchive = []; // Store files to archive after successful processing
  
  try {
    // Step 1: Get template file (don't copy yet)
    const templateFileName = 'Flat.File.ShippingConfirmation.jp.xls';
    const today = new Date();
    const dateStr = today.getFullYear().toString() + 
                   String(today.getMonth() + 1).padStart(2, '0') + 
                   String(today.getDate()).padStart(2, '0');
    const outputFileName = `Amazon出荷通知_${dateStr}.xlsx`;
    
    // Get template file
    const templateFile = getTemplateFile(templateFileName);
    console.log(`テンプレートファイルを見つけました: ${templateFile.getName()}`);
    
    // Step 2: Read CSV files and collect them for archiving
    console.log('CSVファイルを読み込み中...');
    const sagawaData = readCSVFile('佐川.csv');
    const shoshinData = readCSVFile('昭新紙業.csv');
    const fukuyamaData = readCSVFile('福山通運.csv');
    console.log('CSVファイルの読み込みが完了しました。');
    
    // Step 3: Find and read Amazon data file (17-digit filename)
    const amazonDataFile = findAmazonDataFile();
    if (!amazonDataFile) {
      throw new Error('Amazonデータファイルが見つかりません');
    }
    
    console.log(`Amazonデータファイルを検出: ${amazonDataFile.getName()}`);
    console.log(`Amazonファイルサイズ: ${amazonDataFile.getSize()} バイト`);
    
    const amazonData = readAmazonDataFile(amazonDataFile);
    console.log(`Amazonデータを読み込み: ${amazonData.length} 行`);
    
    // Log first few rows for debugging
    if (amazonData.length > 0) {
      console.log('Amazonデータ サンプル（先頭3行）:');
      for (let i = 0; i < Math.min(3, amazonData.length); i++) {
        console.log(`行 ${i}: ${amazonData[i].join(' | ')}`);
      }
    }
    
    // Step 4: Process data and write to Excel using Advanced Drive Service
    if (typeof Drive !== 'undefined') {
      // Use Advanced Drive Service to process Excel with macros
      processExcelViaGoogleSheets(templateFile, amazonData, sagawaData, shoshinData, fukuyamaData, outputFileName);
    } else {
      // Fallback to logging data for manual entry
      processAndWriteData(templateFile, amazonData, sagawaData, shoshinData, fukuyamaData);
    }
    
    console.log('データ処理が正常に完了しました。');
    
    // Step 5: Archive input files after successful processing
    console.log('入力ファイルをアーカイブしています...');
    archiveInputFiles();
    
  } catch (error) {
    console.error('processAmazonShippingData のエラー:', error.message);
    throw error;
  }
}

/**
 * Delete existing files with the same name in the specified folder
 */
function deleteExistingFile(folderId, fileName) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName(fileName);
    
    let deletedCount = 0;
    while (files.hasNext()) {
      const file = files.next();
      file.setTrashed(true);
      deletedCount++;
      console.log(`同名の既存ファイルを削除: ${fileName}`);
    }
    
    if (deletedCount > 0) {
      console.log(`同名の既存ファイルを ${deletedCount} 件削除: ${fileName}`);
    }
    
    return deletedCount;
  } catch (error) {
    console.error(`既存ファイルの削除エラー ${fileName}:`, error.message);
    return 0;
  }
}

/**
 * Copy template file from flatFileFolderId to outputFolderId with new name
 */
function copyTemplateFile(templateFileName, newFileName) {
  try {
    const sourceFolder = DriveApp.getFolderById(flatFileFolderId);
    const files = sourceFolder.getFilesByName(templateFileName);
    
    if (!files.hasNext()) {
      throw new Error(`テンプレートファイル「${templateFileName}」が見つかりません`);
    }
    
    // Delete existing file with the same name in destination folder
    deleteExistingFile(outputFolderId, newFileName);
    
    const sourceFile = files.next();
    const destinationFolder = DriveApp.getFolderById(outputFolderId);
    const copiedFile = sourceFile.makeCopy(newFileName, destinationFolder);
    
    return copiedFile;
  } catch (error) {
    console.error('テンプレートファイルのコピー中にエラー:', error.message);
    throw error;
  }
}

/**
 * Get template file from flatFileFolderId
 */
function getTemplateFile(templateFileName) {
  try {
    const sourceFolder = DriveApp.getFolderById(flatFileFolderId);
    const files = sourceFolder.getFilesByName(templateFileName);
    
    if (!files.hasNext()) {
      throw new Error(`テンプレートファイル「${templateFileName}」が見つかりません`);
    }
    
    const sourceFile = files.next();
    return sourceFile;
  } catch (error) {
    console.error('テンプレートファイルの取得中にエラー:', error.message);
    throw error;
  }
}

/**
 * Read CSV file from inputFolderId
 */
function readCSVFile(fileName) {
  try {
    const folder = DriveApp.getFolderById(inputFolderId);
    const files = folder.getFilesByName(fileName);
    
    if (!files.hasNext()) {
      console.warn(`CSVファイル「${fileName}」が見つかりません`);
      return [];
    }
    
    const file = files.next();
    console.log(`CSV読み込み: ${fileName} (${file.getSize()} バイト)`);
    
    // Try multiple encodings and evaluate which one produces the most readable Japanese text
    const encodings = ['UTF-8', 'Shift_JIS', 'EUC-JP', 'ISO-2022-JP'];
    let bestContent = null;
    let bestEncoding = 'UTF-8';
    let bestScore = -1;
    
    for (const encoding of encodings) {
      try {
        const content = file.getBlob().getDataAsString(encoding);
        
        // Score the content based on Japanese character readability
        const score = evaluateJapaneseReadability(content);
        
        console.log(`CSV「${fileName}」 ${encoding} のエンコード評価: ${score}`);
        
        if (score > bestScore) {
          bestScore = score;
          bestContent = content;
          bestEncoding = encoding;
        }
      } catch (e) {
        console.log(`CSV「${fileName}」 ${encoding} での読み取りに失敗: ${e.message}`);
      }
    }
    
    if (!bestContent) {
      console.error(`CSVファイル「${fileName}」はどのエンコードでも読み取れませんでした`);
      return [];
    }
    
    console.log(`CSV「${fileName}」 内容の文字数: ${bestContent.length}`);
    console.log(`CSV「${fileName}」 最適なエンコード: ${bestEncoding} (スコア: ${bestScore})`);
    console.log(`CSV「${fileName}」 先頭300文字: ${bestContent.substring(0, 300)}`);
    
    const lines = bestContent.split('\n');
    console.log(`CSV「${fileName}」 行数: ${lines.length}`);
    
    // Parse CSV (assuming tab-separated or comma-separated)
    const data = [];
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].trim()) {
        // Try tab-separated first, then comma-separated
        let columns = lines[i].split('\t');
        if (columns.length === 1) {
          columns = lines[i].split(',');
        }
        
        // Clean up columns by removing quotes and trimming whitespace
        const cleanedColumns = columns.map(col => {
          return col.replace(/^["']|["']$/g, '').trim(); // Remove quotes from start and end
        });
        
        console.log(`CSV「${fileName}」 行 ${i + 1}: ${cleanedColumns.length} 列 - ${cleanedColumns.join(' | ')}`);
        data.push(cleanedColumns);
      }
    }
    
    console.log(`CSV「${fileName}」 解析完了: データ行 ${data.length} 行`);
    return data;
  } catch (error) {
    console.error(`CSVファイル読み取りエラー ${fileName}:`, error.message);
    return [];
  }
}

/**
 * Evaluate how readable Japanese text is in a given content
 * Higher score means better readability
 */
function evaluateJapaneseReadability(content) {
  let score = 0;
  
  // Check for common Japanese characters (Hiragana, Katakana, Kanji)
  const japanesePattern = /[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/g;
  const japaneseMatches = content.match(japanesePattern);
  
  if (japaneseMatches) {
    score += japaneseMatches.length * 10; // Bonus for each Japanese character
  }
  
  // Check for broken character patterns (common in encoding issues)
  const brokenPatterns = [
    /[\uFFFD]/g, // Replacement character
    /[\u00A0-\u00FF]/g, // Extended ASCII (often indicates encoding issues)
    /[^\x00-\x7F\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF\s]/g // Non-printable or broken chars
  ];
  
  for (const pattern of brokenPatterns) {
    const matches = content.match(pattern);
    if (matches) {
      score -= matches.length * 5; // Penalty for broken characters
    }
  }
  
  // Bonus for readable Japanese words
  const commonJapaneseWords = ['佐川', '昭新', '福山', '配送', '出荷', '通知', '業者'];
  for (const word of commonJapaneseWords) {
    if (content.includes(word)) {
      score += 50; // Bonus for finding expected Japanese words
    }
  }
  
  // Penalty for excessive garbled text
  const garbledPattern = /[^\x00-\x7F\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF\s]/g;
  const garbledMatches = content.match(garbledPattern);
  if (garbledMatches && garbledMatches.length > content.length * 0.1) {
    score -= 100; // Heavy penalty if more than 10% is garbled
  }
  
  return score;
}

/**
 * Find Amazon data file with 17-digit filename
 */
function findAmazonDataFile() {
  try {
    const folder = DriveApp.getFolderById(inputFolderId);
    const files = folder.getFiles();
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      // Check if filename is 17 digits
      if (/^\d{17}\.txt$/.test(fileName)) {
        return file;
      }
    }
    
    return null;
  } catch (error) {
    console.error('Amazonデータファイルの検索エラー:', error.message);
    return null;
  }
}

/**
 * Read Amazon data file with proper encoding detection
 */
function readAmazonDataFile(file) {
  try {
    // Try multiple encodings and evaluate which one produces the most readable Japanese text
    const encodings = ['UTF-8', 'Shift_JIS', 'EUC-JP', 'ISO-2022-JP'];
    let bestContent = null;
    let bestEncoding = 'UTF-8';
    let bestScore = -1;
    
    for (const encoding of encodings) {
      try {
        const content = file.getBlob().getDataAsString(encoding);
        
        // Score the content based on Japanese character readability
        const score = evaluateJapaneseReadability(content);
        
        console.log(`Amazonファイル ${encoding} のエンコード評価: ${score}`);
        
        if (score > bestScore) {
          bestScore = score;
          bestContent = content;
          bestEncoding = encoding;
        }
      } catch (e) {
        console.log(`Amazonファイル ${encoding} での読み取りに失敗: ${e.message}`);
      }
    }
    
    if (!bestContent) {
      console.error('Amazonファイルはどのエンコードでも読み取れませんでした');
      return [];
    }
    
    console.log(`Amazonファイル 内容の文字数: ${bestContent.length} 文字`);
    console.log(`Amazonファイル 最適なエンコード: ${bestEncoding} (スコア: ${bestScore})`);
    console.log(`Amazonファイル 先頭500文字: ${bestContent.substring(0, 500)}`);
    
    const lines = bestContent.split('\n');
    console.log(`Amazonファイル 行数: ${lines.length}`);
    
    const data = [];
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].trim()) {
        const columns = lines[i].split('\t');
        
        // Clean up columns by removing quotes and trimming whitespace
        const cleanedColumns = columns.map(col => {
          return col.replace(/^["']|["']$/g, '').trim(); // Remove quotes from start and end
        });
        
        console.log(`行 ${i + 1}: ${cleanedColumns.length} 列 - ${cleanedColumns.join(' | ')}`);
        data.push(cleanedColumns);
      }
    }
    
    console.log(`Amazonファイルの解析完了: データ行 ${data.length} 行`);
    return data;
  } catch (error) {
    console.error('Amazonデータファイルの読み取りエラー:', error.message);
    return [];
  }
}

/**
 * Process data and write to Excel template
 * Note: This function cannot directly modify Excel files with macros
 * It will create a log of the data that should be written
 */
function processAndWriteData(templateFile, amazonData, sagawaData, shoshinData, fukuyamaData) {
  try {
    console.log('Amazonデータを処理中...');
    console.log(`テンプレートファイル: ${templateFile.getName()}`);
    console.log(`Amazonデータ行数: ${amazonData.length - 1}`); // Exclude header
    
    // Process each row starting from row 2 of Amazon data
    let outputRow = 4; // Start writing from row 4 in Excel
    const processedData = [];
    
    for (let i = 1; i < amazonData.length; i++) { // Skip header row
      const row = amazonData[i];
      
      if (row.length < 25) continue; // Skip incomplete rows
      
      // Extract data from Amazon file
      const orderId = row[0] || '';
      const orderItemId = row[1] || '';
      //   const purchaseDate = row[2] || '';
      const quantityPurchased = row[9] || '';
      const buyerName = row[16] || '';
      
      // Use current date for shipping date (not purchase date)
      const today = new Date();
      const shippingDate = today.getFullYear().toString() + 
                          '-' + String(today.getMonth() + 1).padStart(2, '0') + 
                          '-' + String(today.getDate()).padStart(2, '0');
      
      // Search for matching data in CSV files
      const matchResult = findMatchingData(buyerName, sagawaData, shoshinData, fukuyamaData);
      
      // Create data row for output
      const outputRowData = {
        row: outputRow,
        orderId: orderId,
        orderItemId: orderItemId,
        quantityPurchased: quantityPurchased,
        convertedDate: shippingDate, // Use current date for shipping
        type: 'Other',
        carrier: matchResult.found ? matchResult.carrier : '',
        trackingNumber: matchResult.found ? matchResult.trackingNumber : ''
      };
      
      processedData.push(outputRowData);
      outputRow++;
    }
    
    // Log the processed data
    console.log('Excelへ書き込む処理済みデータ:');
    console.log('行 | 注文ID | 注文アイテムID | 数量 | 日付 | 種別 | 配送業者 | 追跡番号');
    console.log('----|--------|----------------|------|------|------|----------|----------');
    
    processedData.forEach(data => {
      console.log(`${data.row} | ${data.orderId} | ${data.orderItemId} | ${data.quantityPurchased} | ${data.convertedDate} | ${data.type} | ${data.carrier} | ${data.trackingNumber}`);
    });
    
    console.log(`処理済み行数合計: ${processedData.length}`);
    console.log('注意: マクロ付きExcelはApps Scriptから直接編集できません。');
    console.log('上記データを手動でExcelに転記してください。');
    
  } catch (error) {
    console.error('データ処理/出力中のエラー:', error.message);
    throw error;
  }
}

/**
 * Convert date format from ISO to yyyy-mm-dd
 */
function convertDateFormat(isoDate) {
  try {
    if (!isoDate) return '';
    
    const date = new Date(isoDate);
    if (isNaN(date.getTime())) return '';
    
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    return `${year}-${month}-${day}`;
  } catch (error) {
    console.error('日付形式の変換エラー:', error.message);
    return '';
  }
}

/**
 * Normalize text for comparison (remove spaces, convert to lowercase, etc.)
 */
function normalizeText(text) {
  if (!text) return '';
  
  // Normalize Unicode forms (e.g., full-width/half-width, compatibility chars)
  try {
    if (typeof text.normalize === 'function') {
      text = text.normalize('NFKC');
    }
  } catch (e) {
    // Ignore if normalize is unavailable
  }
  
  return text
    .replace(/\s+/g, '') // Remove all spaces
    .toLowerCase() // Convert to lowercase
    .replace(/[．]/g, '.') // Normalize full-width periods
    .replace(/[　]/g, '') // Remove full-width spaces
    // Normalize middle dots (both full-width and half-width) by removing them
    .replace(/[・･]/g, '')
    // Convert full-width letters to half-width
    .replace(/[Ａ-Ｚ]/g, function(match) {
      return String.fromCharCode(match.charCodeAt(0) - 0xFEE0);
    })
    .replace(/[ａ-ｚ]/g, function(match) {
      return String.fromCharCode(match.charCodeAt(0) - 0xFEE0);
    })
    // Convert full-width numbers to half-width
    .replace(/[０-９]/g, function(match) {
      return String.fromCharCode(match.charCodeAt(0) - 0xFEE0);
    })
    // Convert full-width symbols to half-width
    .replace(/[！＠＃＄％＾＆＊（）＿＋－＝｛｝｜：＂；＇＜＞？，．／]/g, function(match) {
      return String.fromCharCode(match.charCodeAt(0) - 0xFEE0);
    })
    // Remove dash/prolonged sound mark characters entirely (cover many Unicode variants)
    .replace(/[ーｰ゠\-]/g, '') // Katakana prolonged sound marks and ASCII hyphen-minus
    .replace(/[‐‑‒–—―−]/g, '') // U+2010..U+2015 hyphen/dash variants and U+2212 minus sign
    .trim();
}

/**
 * Check if two names match using fuzzy logic
 */
function fuzzyMatch(name1, name2) {
  if (!name1 || !name2) return false;
  
  const normalized1 = normalizeText(name1);
  const normalized2 = normalizeText(name2);
  
  // Exact match after normalization
  if (normalized1 === normalized2) return true;
  
  // Check if one contains the other (for partial matches)
  if (normalized1.includes(normalized2) || normalized2.includes(normalized1)) return true;
  
  // Check for common variations
  const variations1 = [
    normalized1,
    normalized1.replace(/株式会社/g, ''),
    normalized1.replace(/有限会社/g, ''),
    normalized1.replace(/合同会社/g, ''),
    normalized1.replace(/\./g, ''),
    normalized1.replace(/online store/gi, 'オンラインストア'),
    normalized1.replace(/オンラインストア/gi, 'online store')
  ];
  
  const variations2 = [
    normalized2,
    normalized2.replace(/株式会社/g, ''),
    normalized2.replace(/有限会社/g, ''),
    normalized2.replace(/合同会社/g, ''),
    normalized2.replace(/\./g, ''),
    normalized2.replace(/online store/gi, 'オンラインストア'),
    normalized2.replace(/オンラインストア/gi, 'online store')
  ];
  
  // Check if any variations match
  for (const v1 of variations1) {
    for (const v2 of variations2) {
      if (v1 === v2) return true;
      if (v1.includes(v2) || v2.includes(v1)) return true;
    }
  }
  
  return false;
}


/**
 * Get start index for iterating Shoshin data respecting optional header
 */
function getShoshinStartIndex(shoshinData) {
  // if (!shoshinData || shoshinData.length === 0) return 0;
  // return isShoshinHeaderRow(shoshinData[0]) ? 1 : 0;
  return 0;
}

/**
 * Find matching data in CSV files based on buyer name
 */
function findMatchingData(buyerName, sagawaData, shoshinData, fukuyamaData) {
  try {
    console.log(`検索対象の購入者名: "${buyerName}"`);
    
    // Debug: Show some CSV data
    // console.log('佐川CSV サンプル（先頭3行）:');
    // for (let i = 1; i < Math.min(4, sagawaData.length); i++) {
    //   if (sagawaData[i].length > 16) {
    //     console.log(`  行 ${i}: H="${sagawaData[i][7]}" | O="${sagawaData[i][14]}" | P="${sagawaData[i][15]}" | Q="${sagawaData[i][16]}"`);
    //   }
    // }
    
    // console.log('昭新CSV サンプル（先頭3行）:');
    // const shoshinStartIndex = getShoshinStartIndex(shoshinData);
    // for (let i = shoshinStartIndex; i < Math.min(shoshinStartIndex + 3, shoshinData.length); i++) {
    //   if (shoshinData[i].length > 0) {
    //     console.log(`  行 ${i}: "${shoshinData[i][0]}"`);
    //   }
    // }
    
    // console.log('福山CSV サンプル（先頭3行）:');
    // for (let i = 1; i < Math.min(4, fukuyamaData.length); i++) {
    //   if (fukuyamaData[i].length > 9) {
    //     console.log(`  行 ${i}: "${fukuyamaData[i][9]}"`);
    //   }
    // }
    
    // Search in Sagawa data (check multiple columns for buyer names)
    for (let i = 1; i < sagawaData.length; i++) {
      // Check multiple columns where buyer names might be located
      const columnsToCheck = [7, 14, 15, 16]; // Column H, O, P, Q
      
      for (const colIndex of columnsToCheck) {
        if (sagawaData[i].length > colIndex && sagawaData[i][colIndex]) {
          const csvName = sagawaData[i][colIndex];
          if (fuzzyMatch(buyerName, csvName)) {
            console.log(`佐川の列 ${colIndex} で一致を検出: "${buyerName}" ≒ "${csvName}"`);
            return {
              found: true,
              carrier: '佐川急便',
              trackingNumber: sagawaData[i][0] || '' // お問い合せ送り状No.
            };
          }
        }
      }
    }
    
    // Search in Shoshin data (Column A)
    for (let i = shoshinStartIndex; i < shoshinData.length; i++) {
      if (shoshinData[i].length > 0 && shoshinData[i][0]) {
        const csvName = shoshinData[i][0];
        if (fuzzyMatch(buyerName, csvName)) {
          console.log(`昭新で一致を検出: "${buyerName}" ≒ "${csvName}"`);
          return {
            found: true,
            carrier: '佐川急便',
            trackingNumber: shoshinData[i][9] || '' // Column J (index 9) - お問い合わせ伝票番号
          };
        }
      }
    }
    
    // Search in Fukuyama data (Column 9: 荷受人名前１)
    for (let i = 1; i < fukuyamaData.length; i++) {
      if (fukuyamaData[i].length > 9 && fukuyamaData[i][9]) {
        const csvName = fukuyamaData[i][9];
        if (fuzzyMatch(buyerName, csvName)) {
          console.log(`福山で一致を検出: "${buyerName}" ≒ "${csvName}"`);
          return {
            found: true,
            carrier: '福山通運',
            trackingNumber: fukuyamaData[i][2] || '' // 送り状番号
          };
        }
      }
    }
    
    console.log(`一致なし: "${buyerName}"`);
    return { found: false, carrier: '', trackingNumber: '' };
    
  } catch (error) {
    console.error('照合中のエラー:', error.message);
    return { found: false, carrier: '', trackingNumber: '' };
  }
}

/**
 * Utility function to list all files in a folder for debugging
 */
function listFilesInFolder(folderId, folderName) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    
    console.log(`${folderName} (${folderId}) のファイル一覧:`);
    console.log(`フォルダ名: ${folder.getName()}`);
    console.log('---');
    
    while (files.hasNext()) {
      const file = files.next();
      console.log(`名前: ${file.getName()}`);
      console.log(`ID: ${file.getId()}`);
      console.log(`サイズ: ${file.getSize()} バイト`);
      console.log('---');
    }
    
  } catch (error) {
    console.error(`${folderName} の一覧取得エラー:`, error.message);
  }
}

/**
 * Debug function to list all files in all folders
 */
function debugListAllFiles() {
  listFilesInFolder(archiveFolderId, 'アーカイブフォルダ');
  listFilesInFolder(flatFileFolderId, 'テンプレートフォルダ');
  listFilesInFolder(outputFolderId, '出力フォルダ');
  listFilesInFolder(inputFolderId, '入力フォルダ');
}

/**
 * Enable Advanced Drive Service - Run this function first
 */
function enableAdvancedDriveService() {
  try {
    console.log('高度なDriveサービスを有効化する手順:');
    console.log('1. Apps Scriptエディタを開く');
    console.log('2. 左サイドバーの「サービス」をクリック');
    console.log('3. 「+」ボタンをクリック');
    console.log('4. 一覧から「Drive API」を選択');
    console.log('5. 「追加」をクリック');
    console.log('6. プロジェクトを保存');
    console.log('有効化後、Drive.Files.* メソッドが利用できます');
    
    // Test if Drive API is available
    if (typeof Drive !== 'undefined') {
      console.log('高度なDriveサービスは既に有効です。');
    } else {
      console.log('高度なDriveサービスはまだ有効ではありません。');
    }
  } catch (error) {
    console.error('高度なDriveサービス確認中のエラー:', error.message);
  }
}

/**
 * Read Excel file content using Advanced Drive Service
 */
function readExcelFileContent(fileId) {
  try {
    if (typeof Drive === 'undefined') {
      throw new Error('高度なDriveサービスが有効ではありません。先に enableAdvancedDriveService() を実行してください。');
    }
    
    // Get file metadata
    const file = Drive.Files.get(fileId);
    console.log(`Excelファイルを読み込み: ${file.title}`);
    
    // Get file content as binary
    const response = Drive.Files.get(fileId, {alt: 'media'});
    const content = response.getBlob().getDataAsString();
    
    console.log(`ファイル内容の文字数: ${content.length} 文字`);
    return content;
    
  } catch (error) {
    console.error('Excelファイル読み取りのエラー:', error.message);
    throw error;
  }
}

/**
 * Write data to Excel file using Advanced Drive Service
 */
function writeDataToExcelFile(fileId, data) {
  try {
    if (typeof Drive === 'undefined') {
      throw new Error('高度なDriveサービスが有効ではありません。先に enableAdvancedDriveService() を実行してください。');
    }
    
    // Create a blob with the data
    const blob = Utilities.newBlob(data);
    
    // Update the file content
    const response = Drive.Files.update({}, fileId, blob, {
      media: {
        mimeType: 'application/vnd.ms-excel'
      }
    });
    
    console.log(`Excelファイルを更新しました: ${response.title}`);
    return response;
    
  } catch (error) {
    console.error('Excelファイル書き込みのエラー:', error.message);
    throw error;
  }
}

/**
 * Process Excel file with Advanced Drive Service
 */
function processExcelWithAdvancedDriveService(templateFile, amazonData, sagawaData, shoshinData, fukuyamaData) {
  try {
    if (typeof Drive === 'undefined') {
      throw new Error('高度なDriveサービスが有効ではありません。先に enableAdvancedDriveService() を実行してください。');
    }
    
    console.log('高度なDriveサービスでExcelファイルを処理中...');
    
    // Read the Excel file content
    const fileContent = readExcelFileContent(templateFile.getId());
    
    // Process the data
    const processedData = processDataForExcel(amazonData, sagawaData, shoshinData, fukuyamaData);
    
    // Create Excel-compatible data
    const excelData = createExcelData(fileContent, processedData);
    
    // Write back to the file
    writeDataToExcelFile(templateFile.getId(), excelData);
    
    console.log('高度なDriveサービスによるExcel処理が完了しました。');
    
  } catch (error) {
    console.error('高度なDriveサービスでのExcel処理エラー:', error.message);
    throw error;
  }
}

/**
 * Process data for Excel format
 */
function processDataForExcel(amazonData, sagawaData, shoshinData, fukuyamaData) {
  const processedData = [];
  
  for (let i = 1; i < amazonData.length; i++) { // Skip header row
    const row = amazonData[i];
    
    if (row.length < 25) continue; // Skip incomplete rows
    
    // Extract data from Amazon file
    const orderId = row[0] || '';
    const orderItemId = row[1] || '';
    const purchaseDate = row[2] || '';
    const quantityPurchased = row[9] || '';
    const buyerName = row[16] || '';
    
    // Use current date for shipping date (not purchase date)
    const today = new Date();
    const shippingDate = today.getFullYear().toString() + 
                        '-' + String(today.getMonth() + 1).padStart(2, '0') + 
                        '-' + String(today.getDate()).padStart(2, '0');
    
    // Search for matching data in CSV files
    const matchResult = findMatchingData(buyerName, sagawaData, shoshinData, fukuyamaData);
    
    processedData.push({
      orderId: orderId,
      orderItemId: orderItemId,
      quantityPurchased: quantityPurchased,
      convertedDate: shippingDate, // Use current date for shipping
      type: 'Other',
      carrier: matchResult.found ? matchResult.carrier : '',
      trackingNumber: matchResult.found ? matchResult.trackingNumber : ''
    });
  }
  
  return processedData;
}

/**
 * Create Excel data from processed data
 * Note: This is a simplified approach - Excel binary format is complex
 */
function createExcelData(originalContent, processedData) {
  try {
    // This is a simplified approach
    // In practice, you would need to parse and modify the Excel binary format
    // or use a library like SheetJS (xlsx) which isn't available in Google Apps Script
    
    console.log('Excelデータを作成中...');
    console.log(`元データの長さ: ${originalContent.length}`);
    console.log(`処理済みデータ行数: ${processedData.length}`);
    
    // For now, we'll return the original content
    // In a real implementation, you would modify the Excel binary data
    return originalContent;
    
  } catch (error) {
    console.error('Excelデータ作成中のエラー:', error.message);
    throw error;
  }
}

/**
 * Alternative: Use Google Sheets as intermediate step
 */
function processExcelViaGoogleSheets(templateFile, amazonData, sagawaData, shoshinData, fukuyamaData, outputFileName) {
  try {
    if (typeof Drive === 'undefined') {
      throw new Error('高度なDriveサービスが有効ではありません。先に enableAdvancedDriveService() を実行してください。');
    }
    
    console.log('Googleスプレッドシート経由でExcelを処理中...');
    
    // Step 1: Convert Excel to Google Sheets
    const spreadsheet = convertExcelToGoogleSheets(templateFile);
    
    // Step 2: Process data in Google Sheets
    processDataInGoogleSheets(spreadsheet, amazonData, sagawaData, shoshinData, fukuyamaData);
    
    // Step 3: Export back to Excel
    const excelFile = exportGoogleSheetsToExcel(spreadsheet, outputFileName);
    
    console.log('Googleスプレッドシート経由のExcel処理が完了しました。');
    return excelFile;
    
  } catch (error) {
    console.error('Googleスプレッドシート経由の処理エラー:', error.message);
    throw error;
  }
}

/**
 * Convert Excel to Google Sheets using Drive API
 */
function convertExcelToGoogleSheets(excelFile) {
  try {
    const blob = excelFile.getBlob();
    const resource = {
      title: excelFile.getName() + '_converted',
      mimeType: 'application/vnd.google-apps.spreadsheet'
    };
    
    const convertedFile = Drive.Files.insert(resource, blob, {
      convert: true
    });
    
    const spreadsheet = SpreadsheetApp.openById(convertedFile.id);
    console.log(`ExcelをGoogleスプレッドシートに変換しました: ${convertedFile.title}`);
    
    return spreadsheet;
    
  } catch (error) {
    console.error('Excelからスプレッドシートへの変換エラー:', error.message);
    throw error;
  }
}

/**
 * Process data in Google Sheets
 */
function processDataInGoogleSheets(spreadsheet, amazonData, sagawaData, shoshinData, fukuyamaData) {
  try {
    console.log(`スプレッドシートでデータを処理中: ${spreadsheet.getName()}`);
    console.log(`Amazonデータ行数: ${amazonData.length}`);
    console.log(`佐川データ行数: ${sagawaData.length}`);
    console.log(`昭新データ行数: ${shoshinData.length}`);
    console.log(`福山データ行数: ${fukuyamaData.length}`);
    
    // List all available sheets
    const sheets = spreadsheet.getSheets();
    console.log('利用可能なシート:');
    sheets.forEach((sheet, index) => {
      console.log(`  ${index + 1}. ${sheet.getName()}`);
    });
    
    let sheet = spreadsheet.getSheetByName('出荷通知テンプレート_Template');
    
    if (!sheet) {
      console.log('テンプレートシートが見つかりません。別名を試します...');
      // Try alternative sheet names
      const alternativeNames = ['Template', 'テンプレート', 'Sheet1', 'Sheet 1'];
      for (const name of alternativeNames) {
        sheet = spreadsheet.getSheetByName(name);
        if (sheet) {
          console.log(`シートを使用します: ${name}`);
          break;
        }
      }
      
      if (!sheet) {
        console.log('テンプレートシートがないため、先頭のシートを使用します');
        sheet = spreadsheet.getSheets()[0];
      }
    }
    
    console.log(`使用シート: ${sheet.getName()}`);
    console.log(`シートサイズ: ${sheet.getLastRow()} 行 x ${sheet.getLastColumn()} 列`);
    
    // Clear existing data from row 4 onwards
    const lastRow = sheet.getLastRow();
    if (lastRow >= 4) {
      sheet.getRange(4, 1, lastRow - 3, sheet.getLastColumn()).clear();
      console.log(`行 4-${lastRow} をクリアしました`);
    }
    
    // Write processed data
    let outputRow = 4;
    let processedCount = 0;
    
    for (let i = 1; i < amazonData.length; i++) {
      const row = amazonData[i];
      if (row.length < 25) {
        console.log(`行 ${i} をスキップ: 列数不足 (${row.length})`);
        continue;
      }
      
      const orderId = row[0] || '';
      const orderItemId = row[1] || '';
      //   const purchaseDate = row[2] || '';
      const quantityPurchased = row[9] || '';
      const buyerName = row[16] || ''; // Column Q (index 16) instead of F (index 5)
      
      // Use current date for shipping date (not purchase date)
      const today = new Date();
      const shippingDate = today.getFullYear().toString() + 
                          '-' + String(today.getMonth() + 1).padStart(2, '0') + 
                          '-' + String(today.getDate()).padStart(2, '0');
      
      const matchResult = findMatchingData(buyerName, sagawaData, shoshinData, fukuyamaData);
      
      console.log(`処理中 ${i}: 注文ID=${orderId}, 購入者=${buyerName}, 照合=${matchResult.found}`);
      
      // Write data to specific cells
      sheet.getRange(outputRow, 1).setValue(orderId);
      sheet.getRange(outputRow, 2).setValue(orderItemId);
      sheet.getRange(outputRow, 3).setValue(quantityPurchased);
      sheet.getRange(outputRow, 4).setValue(shippingDate); // Use current date for shipping
      sheet.getRange(outputRow, 5).setValue('Other');
      
      if (matchResult.found) {
        sheet.getRange(outputRow, 6).setValue(matchResult.carrier);
        sheet.getRange(outputRow, 7).setValue(matchResult.trackingNumber);
      }
      
      // Verify data was written
      const writtenData = sheet.getRange(outputRow, 1, 1, 7).getValues()[0];
      console.log(`行 ${outputRow} に書き込み: ${writtenData.join(' | ')}`);
      
      outputRow++;
      processedCount++;
    }
    
    console.log(`Googleスプレッドシートへの書き込み完了。処理件数: ${processedCount}`);
    
  } catch (error) {
    console.error('スプレッドシート処理中のエラー:', error.message);
    throw error;
  }
}

/**
 * Export Google Sheets to Excel
 */
function exportGoogleSheetsToExcel(spreadsheet, fileName) {
  try {
    console.log(`エクスポート対象: ${spreadsheet.getName()}`);
    console.log(`スプレッドシートID: ${spreadsheet.getId()}`);
    
    // Delete existing file with the same name in destination folder
    deleteExistingFile(outputFolderId, fileName);
    
    // Check data before export
    const sheets = spreadsheet.getSheets();
    console.log('スプレッドシート内のシート一覧:');
    sheets.forEach((sheet, index) => {
      console.log(`  ${index + 1}. ${sheet.getName()} - ${sheet.getLastRow()} 行 x ${sheet.getLastColumn()} 列`);
      if (sheet.getName() === '出荷通知テンプレート_Template') {
        const dataRange = sheet.getRange(4, 1, sheet.getLastRow() - 3, 7);
        const data = dataRange.getValues();
        console.log('  テンプレートシートのデータ（先頭3行）:');
        for (let i = 0; i < Math.min(3, data.length); i++) {
          console.log(`    行 ${i + 4}: ${data[i].join(' | ')}`);
        }
      }
    });
    
    // Export Google Sheets as Excel using the proper export URL
    const spreadsheetId = spreadsheet.getId();
    const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx`;
    
    console.log(`エクスポートURL: ${exportUrl}`);
    
    // Get the exported Excel file as a blob
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });
    
    const excelBlob = response.getBlob().setName(fileName);
    const destinationFolder = DriveApp.getFolderById(outputFolderId);
    const excelFile = destinationFolder.createFile(excelBlob);
    
    // Clean up temporary spreadsheet
    DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
    
    console.log(`Excelとしてエクスポートしました: ${excelFile.getName()}`);
    console.log(`Excelファイルサイズ: ${excelFile.getSize()} バイト`);
    return excelFile;
    
  } catch (error) {
    console.error('Excelへのエクスポートエラー:', error.message);
    throw error;
  }
}

/**
 * Monitor and detect file uploads in inputFolderId
 * Returns information about required files and their status
 */
function detectFileUploads() {
  try {
    console.log('入力フォルダ内の必要ファイルを確認しています...');
    
    const folder = DriveApp.getFolderById(inputFolderId);
    const files = folder.getFiles();
    
    const requiredFiles = {
      '佐川.csv': { found: false, file: null, size: 0, lastModified: null },
      '昭新紙業.csv': { found: false, file: null, size: 0, lastModified: null },
      '福山通運.csv': { found: false, file: null, size: 0, lastModified: null },
      'amazon_txt': { found: false, file: null, size: 0, lastModified: null, pattern: /^\d{17}\.txt$/ }
    };
    
    let allFiles = [];
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const fileSize = file.getSize();
      const lastModified = file.getLastUpdated();
      
      allFiles.push({
        name: fileName,
        size: fileSize,
        lastModified: lastModified,
        id: file.getId()
      });
      
      // Check for required CSV files
      if (requiredFiles.hasOwnProperty(fileName)) {
        requiredFiles[fileName].found = true;
        requiredFiles[fileName].file = file;
        requiredFiles[fileName].size = fileSize;
        requiredFiles[fileName].lastModified = lastModified;
        console.log(`✓ 必要ファイルを検出: ${fileName} (${fileSize} バイト)`);
      }
      
      // Check for Amazon TXT file (17-digit filename)
      if (requiredFiles.amazon_txt.pattern.test(fileName)) {
        requiredFiles.amazon_txt.found = true;
        requiredFiles.amazon_txt.file = file;
        requiredFiles.amazon_txt.size = fileSize;
        requiredFiles.amazon_txt.lastModified = lastModified;
        console.log(`✓ AmazonのTXTファイルを検出: ${fileName} (${fileSize} バイト)`);
      }
    }
    
    // Report status
    console.log('\n=== ファイルアップロード状況 ===');
    let allRequiredFilesFound = true;
    
    for (const [fileName, status] of Object.entries(requiredFiles)) {
      if (fileName === 'amazon_txt') {
        if (status.found) {
          console.log(`✓ Amazon TXTファイル: ${status.file.getName()} (${status.size} バイト)`);
        } else {
          console.log('✗ Amazon TXTファイル: 見つかりません（例: 12345678901234567.txt のような17桁のファイル名）');
          allRequiredFilesFound = false;
        }
      } else {
        if (status.found) {
          console.log(`✓ ${fileName}: 検出 (${status.size} バイト)`);
        } else {
          console.log(`✗ ${fileName}: 見つかりません`);
          allRequiredFilesFound = false;
        }
      }
    }
    
    console.log('\n=== 入力フォルダ内の全ファイル ===');
    allFiles.forEach(file => {
      const dateStr = file.lastModified.toLocaleString();
      console.log(`- ${file.name} (${file.size} バイト, 更新日: ${dateStr})`);
    });
    
    if (allRequiredFilesFound) {
      console.log('\n✅ 必要なファイルがすべて揃いました。処理可能です。');
      return {
        ready: true,
        files: requiredFiles,
        allFiles: allFiles
      };
    } else {
      console.log('\n❌ 必要なファイルが不足しています。不足分をアップロードしてください。');
      return {
        ready: false,
        files: requiredFiles,
        allFiles: allFiles
      };
    }
    
  } catch (error) {
    console.error('ファイルアップロード検知のエラー:', error.message);
    return {
      ready: false,
      error: error.message,
      files: {},
      allFiles: []
    };
  }
}

/**
 * Auto-process when all required files are detected
 */
function autoProcessOnUpload() {
  try {
    console.log('アップロードを確認し、自動処理を実行します...');
    
    const uploadStatus = detectFileUploads();
    
    if (uploadStatus.ready) {
      console.log('必要なファイルを検出。自動処理を開始します...');
      
      // Check if files are recent (within last hour)
      const now = new Date();
      const oneHourAgo = new Date(now.getTime() - 60 * 60 * 1000);
      
      let hasRecentFiles = false;
      for (const [fileName, status] of Object.entries(uploadStatus.files)) {
        if (status.found && status.lastModified && status.lastModified > oneHourAgo) {
          hasRecentFiles = true;
          break;
        }
      }
      
      if (hasRecentFiles) {
        console.log('直近1時間以内のアップロードを検出。自動処理を実行します...');
        processAmazonShippingData();
      } else {
        console.log('ファイルは存在しますが更新が古いため、自動処理は行いません。手動実行してください。');
      }
    } else {
      console.log('必要なファイルが不足しているため、自動処理できません。');
    }
    
  } catch (error) {
    console.error('自動処理のエラー:', error.message);
  }
}

/**
 * Set up time-based trigger to check for file uploads
 * Note: This only checks at scheduled intervals, not on actual file uploads
 */
function setupFileUploadTrigger() {
  try {
    // Delete existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'checkAndProcessFiles') {
        ScriptApp.deleteTrigger(trigger);
        console.log('既存のトリガーを削除しました');
      }
    });
    
    // Create new trigger to run every 5 minutes
    ScriptApp.newTrigger('checkAndProcessFiles')
      .timeBased()
      .everyMinutes(5)
      .create();
    
    console.log('✅ ファイルアップロード監視トリガーを作成しました');
    console.log('📅 5分ごとにチェックします');
    console.log('⚠️  これはイベントではなく時間ベースのチェックです');
    console.log('⚠️  最小間隔は5分です（1分は不可）');
    
  } catch (error) {
    console.error('トリガー設定中のエラー:', error.message);
  }
}

/**
 * Check for files and process if all are present
 * This function is called by the time-based trigger
 */
function checkAndProcessFiles() {
  try {
    console.log('🕐 定期実行: ファイルアップロードの確認...');
    
    const uploadStatus = detectFileUploads();
    
    if (uploadStatus.ready) {
      console.log('✅ 必要なファイルを検出。処理を開始します...');
      processAmazonShippingData();
    } else {
      console.log('❌ 必要なファイルが不足。アップロードを待機します...');
    }
    
  } catch (error) {
    console.error('定期チェック中のエラー:', error.message);
  }
}

/**
 * List all current triggers
 */
function listTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    
    console.log('📋 現在のトリガー一覧:');
    if (triggers.length === 0) {
      console.log('トリガーはありません');
    } else {
      triggers.forEach((trigger, index) => {
        console.log(`${index + 1}. 関数: ${trigger.getHandlerFunction()}`);
        console.log(`   種別: ${trigger.getEventType()}`);
        console.log(`   ID: ${trigger.getUniqueId()}`);
        console.log('---');
      });
    }
    
  } catch (error) {
    console.error('トリガー一覧取得中のエラー:', error.message);
  }
}

/**
 * Delete all triggers
 */
function deleteAllTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    
    triggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      console.log(`トリガーを削除: ${trigger.getHandlerFunction()}`);
    });
    
    console.log('✅ すべてのトリガーを削除しました');
    
  } catch (error) {
    console.error('トリガー削除中のエラー:', error.message);
  }
}

/**
 * Archive input files to the archive folder with force removal
 */
function archiveInputFiles() {
  try {
    const inputFolder = DriveApp.getFolderById(inputFolderId);
    const archiveFolder = DriveApp.getFolderById(archiveFolderId);
    
    // Files to archive
    const filesToArchive = [
      '佐川.csv',
      '昭新紙業.csv', 
      '福山通運.csv'
    ];
    
    // Also find and archive Amazon data file (17-digit filename)
    const amazonFile = findAmazonDataFile();
    if (amazonFile) {
      filesToArchive.push(amazonFile.getName());
    }
    
    let archivedCount = 0;
    let removedCount = 0;
    
    for (const fileName of filesToArchive) {
      try {
        const files = inputFolder.getFilesByName(fileName);
        
        while (files.hasNext()) {
          const file = files.next();
          
          try {
            // Create a copy in the archive folder with date format YYYYMMDD
            const today = new Date();
            const dateStr = today.getFullYear().toString() + 
                           String(today.getMonth() + 1).padStart(2, '0') + 
                           String(today.getDate()).padStart(2, '0');
            
            const fileExtension = fileName.match(/\.[^/.]+$/)[0];
            const fileNameWithoutExtension = fileName.replace(/\.[^/.]+$/, '');
            const archivedFileName = `${fileNameWithoutExtension}_${dateStr}${fileExtension}`;
            
            // Try to copy file first
            try {
              const archivedFile = file.makeCopy(archivedFileName, archiveFolder);
              console.log(`アーカイブ: ${fileName} -> ${archivedFileName}`);
              archivedCount++;
            } catch (copyError) {
              console.log(`コピー失敗 ${fileName}: ${copyError.message}`);
              // Continue with removal even if copy fails
            }
            
            // Force remove the original file using multiple methods
            let fileRemoved = false;
            
            // Method 1: Try to move to trash
            try {
              file.setTrashed(true);
              console.log(`ゴミ箱へ移動: ${fileName}`);
              fileRemoved = true;
            } catch (trashError) {
              console.log(`ゴミ箱移動に失敗: ${trashError.message}`);
            }
            
            // Method 2: Try to delete permanently if trash failed
            if (!fileRemoved) {
              try {
                file.setTrashed(false); // First un-trash if it was already trashed
                file.setTrashed(true);  // Then trash again
                console.log(`強制的にゴミ箱へ移動: ${fileName}`);
                fileRemoved = true;
              } catch (forceTrashError) {
                console.log(`強制ゴミ箱移動に失敗: ${forceTrashError.message}`);
              }
            }
            
            // Method 3: Try to remove from folder (if we have folder permissions)
            if (!fileRemoved) {
              try {
                // Get all parents and try to remove from each
                const parents = file.getParents();
                while (parents.hasNext()) {
                  const parent = parents.next();
                  try {
                    parent.removeFile(file);
                    console.log(`フォルダから削除: ${fileName}`);
                    fileRemoved = true;
                    break;
                  } catch (removeError) {
                    console.log(`フォルダからの削除に失敗: ${removeError.message}`);
                  }
                }
              } catch (parentError) {
                console.log(`親フォルダ取得に失敗: ${parentError.message}`);
              }
            }
            
            // Method 4: Try using Drive API directly (if available)
            if (!fileRemoved && typeof Drive !== 'undefined') {
              try {
                Drive.Files.remove(file.getId());
                console.log(`Drive APIで完全削除: ${fileName}`);
                fileRemoved = true;
              } catch (driveApiError) {
                console.log(`Drive APIによる削除に失敗: ${driveApiError.message}`);
              }
            }
            
            if (fileRemoved) {
              removedCount++;
            } else {
              console.log(`⚠️  ${fileName} を削除できませんでした。手動での削除が必要です。`);
            }
            
          } catch (fileError) {
            console.error(`ファイル処理中のエラー ${fileName}:`, fileError.message);
          }
        }
      } catch (error) {
        console.error(`アーカイブ処理エラー ${fileName}:`, error.message);
      }
    }
    
    console.log(`アーカイブフォルダへ ${archivedCount} 件を保存しました`);
    console.log(`元ファイルを ${removedCount} 件削除/移動しました`);
    
    if (removedCount < archivedCount) {
      console.log('⚠️  アーカイブは完了しましたが、一部の元ファイルを削除できませんでした。権限や共有ドライブの設定をご確認ください。');
    }
    
  } catch (error) {
    console.error('archiveInputFiles のエラー:', error.message);
    throw error;
  }
}

/**
 * Quick test for the specific business name matching issue
 */
function debugMatching() {
  const amazonName = 'CL1松原天美我堂・ショートステイ';
  const fukuyamaName = 'ＣＬ１松原天美我堂・ショ－トステイ';
  
  console.log('=== マッチングデバッグ ===');
  console.log(`Amazon: "${amazonName}"`);
  console.log(`福山: "${fukuyamaName}"`);
  
  const norm1 = normalizeText(amazonName);
  const norm2 = normalizeText(fukuyamaName);
  
  console.log(`正規化後Amazon: "${norm1}"`);
  console.log(`正規化後福山: "${norm2}"`);
  console.log(`一致: ${norm1 === norm2}`);
  console.log(`Amazon包含福山: ${norm1.includes(norm2)}`);
  console.log(`福山包含Amazon: ${norm2.includes(norm1)}`);
  
  const match = fuzzyMatch(amazonName, fukuyamaName);
  console.log(`fuzzyMatch結果: ${match}`);
}
