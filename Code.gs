/**
 * Advanced URL to Google Drive Uploader with Web App UI
 */

// این تابع محیط گرافیکی (HTML) را اجرا می‌کند
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('🚀 Drive Uploader Pro')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// سیستم جمع‌آوری لاگ برای ارسال به محیط کاربری
let executionLogs = [];
function logMsg(msg) {
  console.log(msg);
  executionLogs.push(msg);
}

/**
 * تابع اصلی که از طریق محیط گرافیکی فراخوانی می‌شود
 */
function processUploadFromUI(url, folderName) {
  executionLogs = []; // پاکسازی لاگ‌های قبلی
  
  try {
    logMsg(`⏳ [START] شروع پردازش درخواست آپلود...`);
    logMsg(`🔗 لینک: ${url}`);
    
    // ۱. مدیریت پوشه
    let folderId = null;
    if (folderName && folderName.trim() !== "") {
      folderId = getOrCreateFolder(folderName.trim());
    }

    // ۲. دریافت اطلاعات فایل
    const metadata = fetchFileMetadata(url);
    if (!metadata) {
      throw new Error("سرور مبدا اجازه دسترسی به اطلاعات فایل را نمی‌دهد یا لینک نامعتبر است.");
    }
    
    logMsg(`📊 [INFO] اطلاعات فایل استخراج شد:`);
    logMsg(`   - نام: ${metadata.name}`);
    logMsg(`   - حجم: ${formatBytes(metadata.size)}`);
    logMsg(`   - قابلیت تکه‌تکه (Range): ${metadata.supportsRange ? 'دارد ✅' : 'ندارد ❌'}`);

    // ۳. تصمیم‌گیری برای نحوه آپلود
    const CHUNK_SIZE = 20 * 1024 * 1024; // 20 MB
    let fileId = "";

    if (metadata.size > 0 && metadata.size <= 40 * 1024 * 1024) {
      logMsg("⚡ [INFO] حجم فایل مناسب است. استفاده از روش آپلود یکپارچه و مستقیم...");
      fileId = directUpload(url, folderId, metadata);
    } else if (metadata.supportsRange && metadata.size > 0) {
      logMsg("🔄 [INFO] حجم فایل بالاست. آغاز آپلود پیشرفته تکه‌تکه (Resumable)...");
      fileId = resumableUpload(url, folderId, metadata, CHUNK_SIZE);
    } else {
      logMsg("⚠️ [WARN] سرور از دانلود تکه‌تکه پشتیبانی نمی‌کند. تلاش برای آپلود مستقیم...");
      fileId = directUpload(url, folderId, metadata);
    }

    logMsg(`🎉 [SUCCESS] عملیات با موفقیت ۱۰۰٪ به پایان رسید.`);
    return { success: true, fileId: fileId, fileName: metadata.name, logs: executionLogs };

  } catch (error) {
    logMsg(`❌ [ERROR] ${error.message}`);
    return { success: false, error: error.message, logs: executionLogs };
  }
}

/**
 * جستجو یا ساخت پوشه جدید
 */
function getOrCreateFolder(folderName) {
  logMsg(`📂 [INFO] در حال بررسی پوشه هدف: "${folderName}"`);
  const folders = DriveApp.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    const existingFolder = folders.next();
    logMsg(`✅ [INFO] پوشه "${folderName}" از قبل وجود داشت.`);
    return existingFolder.getId();
  } else {
    logMsg(`🏗️ [INFO] پوشه پیدا نشد. در حال ایجاد پوشه جدید...`);
    const newFolder = DriveApp.getRootFolder().createFolder(folderName);
    logMsg(`✅ [INFO] پوشه با موفقیت ساخته شد.`);
    return newFolder.getId();
  }
}

/**
 * دریافت اطلاعات فراداده (نسخه اصلاح شده با GET)
 */
function fetchFileMetadata(url) {
  try {
    const options = {
      method: "get", 
      headers: { "Range": "bytes=0-1" },
      muteHttpExceptions: true,
      followRedirects: true
    };
    
    let response = UrlFetchApp.fetch(url, options);
    let headers = response.getHeaders();
    
    let fileName = "Downloaded_File_" + Date.now();
    const contentDisposition = headers["Content-Disposition"] || headers["content-disposition"];
    if (contentDisposition && contentDisposition.includes("filename=")) {
      let match = contentDisposition.match(/filename="?([^";]+)"?/);
      if (match && match[1]) fileName = match[1];
    } else {
      let urlPath = url.split('?')[0].split('/');
      let lastSegment = urlPath[urlPath.length - 1];
      if (lastSegment) fileName = decodeURIComponent(lastSegment);
    }

    let size = 0;
    const contentRange = headers["Content-Range"] || headers["content-range"];
    if (contentRange) {
      let match = contentRange.match(/\/(\d+)$/);
      if (match && match[1]) size = parseInt(match[1]);
    }
    
    if (size === 0) {
      const contentLength = headers["Content-Length"] || headers["content-length"];
      if (contentLength) size = parseInt(contentLength);
    }

    const code = response.getResponseCode();
    const supportsRange = code === 206 || (headers["Accept-Ranges"] || headers["accept-ranges"]) === "bytes";
    const mimeType = headers["Content-Type"] || headers["content-type"] || "application/octet-stream";

    if (code >= 400) return null;

    return { name: fileName, size: size, mimeType: mimeType, supportsRange: supportsRange };
  } catch (e) {
    logMsg(`خطا در Metadata: ${e.message}`);
    return null;
  }
}

/**
 * آپلود تکه تکه (Resumable)
 */
function resumableUpload(url, folderId, metadata, chunkSize) {
  const token = ScriptApp.getOAuthToken();
  const uploadUrl = "https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable";
  
  const fileMetadata = { name: metadata.name, mimeType: metadata.mimeType };
  if (folderId) fileMetadata.parents = [folderId];

  const initOptions = {
    method: "post",
    headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
    payload: JSON.stringify(fileMetadata),
    muteHttpExceptions: true
  };
  
  const initResponse = UrlFetchApp.fetch(uploadUrl, initOptions);
  if (initResponse.getResponseCode() !== 200) throw new Error("نشست گوگل درایو ایجاد نشد.");
  
  let headers = initResponse.getHeaders();
  const sessionUri = headers["Location"] || headers["location"];
  
  let start = 0;
  while (start < metadata.size) {
    let end = Math.min(start + chunkSize - 1, metadata.size - 1);
    let percentage = Math.round((end / metadata.size) * 100);
    logMsg(`📦 انتقال تکه: ${formatBytes(start)} تا ${formatBytes(end)} (${percentage}%) ...`);
    
    let chunkResponse = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "Range": `bytes=${start}-${end}` },
      muteHttpExceptions: true,
      followRedirects: true
    });
    
    if (chunkResponse.getResponseCode() >= 400) throw new Error("خطا در دانلود از سرور مبدا.");
    
    let putResponse = UrlFetchApp.fetch(sessionUri, {
      method: "put",
      headers: { "Content-Range": `bytes ${start}-${end}/${metadata.size}` },
      payload: chunkResponse.getBlob(),
      muteHttpExceptions: true
    });
    
    let putCode = putResponse.getResponseCode();
    if (putCode === 200 || putCode === 201) {
      let result = JSON.parse(putResponse.getContentText());
      return result.id;
    } else if (putCode !== 308) {
      throw new Error(`خطای نامشخص آپلود: ${putCode}`);
    }
    start += chunkSize;
  }
}

/**
 * آپلود مستقیم فایل‌های کوچک
 */
function directUpload(url, folderId, metadata) {
  logMsg(`⬇️ در حال دانلود کل فایل در حافظه...`);
  const response = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true, followRedirects: true });
  
  if (response.getResponseCode() >= 400) throw new Error("دانلود فایل شکست خورد.");
  
  let blob = response.getBlob();
  blob.setName(metadata.name);
  
  logMsg(`⬆️ در حال ذخیره در گوگل درایو...`);
  let folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  let file = folder.createFile(blob);
  
  return file.getId();
}

/**
 * فرمت حجم
 */
function formatBytes(bytes) {
  if (!bytes || bytes === 0) return '0 Bytes';
  const k = 1024, sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}
