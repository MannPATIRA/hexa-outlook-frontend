/**
 * Attachment Utility
 * Handles fetching and encoding PDF files for email attachments
 * Supports both Graph API (base64) and Office.js (URL or base64) formats
 */

const AttachmentUtils = {
    // List of default PDF attachments to include with RFQs
    DEFAULT_ATTACHMENTS: [
        'RFQ_Template.pdf',
        'Terms_Conditions.pdf'
    ],

    /**
     * Get the base URL for assets
     * Tries to determine from current location, falls back to Vercel URL
     */
    getAssetsBaseUrl() {
        // Try to get from current window location
        if (typeof window !== 'undefined' && window.location) {
            const origin = window.location.origin;
            const pathname = window.location.pathname;
            
            console.log(`[AttachmentUtils] Current location: ${origin}${pathname}`);
            
            // Assets are always at the root level: /assets/attachments/
            // If pathname is /src/taskpane/taskpane.html, we need to go to root
            // If pathname is /taskpane/taskpane.html, we need to go to root
            // Always use root-level assets
            const rootUrl = origin + '/assets/attachments/';
            console.log(`[AttachmentUtils] Using root-level assets URL: ${rootUrl}`);
            return rootUrl;
        }
        
        // Fallback to Vercel deployment URL
        const fallbackUrl = 'https://hexa-outlook-frontend.vercel.app/assets/attachments/';
        console.log(`[AttachmentUtils] Using fallback URL: ${fallbackUrl}`);
        return fallbackUrl;
    },

    /**
     * Fetch a PDF file and convert to base64
     * @param {string} filename - Name of the PDF file
     * @returns {Promise<string>} Base64-encoded content
     */
    async fetchPdfAsBase64(filename) {
        try {
            const url = this.getAssetsBaseUrl() + filename;
            console.log(`[AttachmentUtils] Fetching PDF from: ${url}`);
            
            const response = await fetch(url, {
                method: 'GET',
                cache: 'no-cache'
            });
            
            if (!response.ok) {
                const errorText = await response.text().catch(() => '');
                throw new Error(`Failed to fetch ${filename}: ${response.status} ${response.statusText}. ${errorText}`);
            }
            
            const contentType = response.headers.get('content-type');
            console.log(`[AttachmentUtils] Response content-type: ${contentType}`);
            
            const blob = await response.blob();
            console.log(`[AttachmentUtils] Blob size: ${blob.size} bytes`);
            
            if (blob.size === 0) {
                throw new Error(`PDF file ${filename} is empty`);
            }
            
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onloadend = () => {
                    try {
                        const dataUrl = reader.result;
                        if (!dataUrl || typeof dataUrl !== 'string') {
                            reject(new Error(`Failed to read PDF ${filename} as data URL`));
                            return;
                        }
                        
                        // Remove data URL prefix (data:application/pdf;base64,)
                        const parts = dataUrl.split(',');
                        if (parts.length < 2) {
                            reject(new Error(`Invalid data URL format for ${filename}`));
                            return;
                        }
                        
                        const base64 = parts[1];
                        if (!base64 || base64.length === 0) {
                            reject(new Error(`Empty base64 content for ${filename}`));
                            return;
                        }
                        
                        console.log(`[AttachmentUtils] ✓ Successfully encoded ${filename} (${base64.length} base64 chars)`);
                        resolve(base64);
                    } catch (err) {
                        reject(err);
                    }
                };
                reader.onerror = (error) => {
                    console.error(`[AttachmentUtils] FileReader error for ${filename}:`, error);
                    reject(new Error(`Failed to read PDF ${filename}: ${error.message || 'Unknown error'}`));
                };
                reader.readAsDataURL(blob);
            });
        } catch (error) {
            console.error(`[AttachmentUtils] ✗ Error fetching PDF ${filename}:`, error);
            console.error(`[AttachmentUtils] Error details:`, {
                message: error.message,
                stack: error.stack,
                url: this.getAssetsBaseUrl() + filename
            });
            throw error;
        }
    },

    /**
     * Get attachment URL for Office.js format
     * @param {string} filename - Name of the PDF file
     * @returns {string} Full URL to the PDF file
     */
    getAttachmentUrl(filename) {
        return this.getAssetsBaseUrl() + filename;
    },

    /**
     * Prepare attachments for Graph API format
     * @param {Array<string>} filenames - Array of PDF filenames
     * @returns {Promise<Array>} Array of Graph API attachment objects
     */
    async prepareGraphApiAttachments(filenames = this.DEFAULT_ATTACHMENTS) {
        const attachments = [];
        
        console.log(`[AttachmentUtils] Preparing ${filenames.length} attachment(s) for Graph API...`);
        
        for (const filename of filenames) {
            try {
                console.log(`[AttachmentUtils] Processing ${filename}...`);
                const contentBytes = await this.fetchPdfAsBase64(filename);
                
                if (!contentBytes || contentBytes.length === 0) {
                    throw new Error(`Empty base64 content for ${filename}`);
                }
                
                const attachment = {
                    '@odata.type': '#microsoft.graph.fileAttachment',
                    name: filename,
                    contentType: 'application/pdf',
                    contentBytes: contentBytes
                };
                
                attachments.push(attachment);
                console.log(`[AttachmentUtils] ✓ Prepared Graph API attachment: ${filename} (${contentBytes.length} base64 chars)`);
            } catch (error) {
                console.error(`[AttachmentUtils] ✗ Failed to prepare attachment ${filename}:`, error);
                console.error(`[AttachmentUtils] Error details:`, {
                    message: error.message,
                    stack: error.stack
                });
                // Continue with other attachments even if one fails
            }
        }
        
        console.log(`[AttachmentUtils] Total attachments prepared: ${attachments.length} out of ${filenames.length}`);
        
        if (attachments.length === 0) {
            throw new Error('No attachments could be prepared. Check console for details.');
        }
        
        return attachments;
    },

    /**
     * Prepare attachments for Office.js format
     * @param {Array<string>} filenames - Array of PDF filenames
     * @param {boolean} useBase64 - If true, use base64; if false, use URLs
     * @returns {Promise<Array>} Array of Office.js attachment objects
     */
    async prepareOfficeJsAttachments(filenames = this.DEFAULT_ATTACHMENTS, useBase64 = false) {
        const attachments = [];
        
        for (const filename of filenames) {
            try {
                if (useBase64) {
                    // Use base64 for Office.js (more reliable but larger)
                    const contentBytes = await this.fetchPdfAsBase64(filename);
                    attachments.push({
                        type: 'file',
                        name: filename,
                        content: contentBytes
                    });
                } else {
                    // Use URL (simpler but requires file to be accessible)
                    attachments.push({
                        type: 'file',
                        name: filename,
                        url: this.getAttachmentUrl(filename)
                    });
                }
                console.log(`✓ Prepared Office.js attachment: ${filename}`);
            } catch (error) {
                console.warn(`⚠ Failed to prepare attachment ${filename}:`, error.message);
                // Continue with other attachments even if one fails
            }
        }
        
        return attachments;
    },

    /**
     * Prepare attachments in both formats (for maximum compatibility)
     * @param {Array<string>} filenames - Array of PDF filenames
     * @returns {Promise<Object>} Object with graphApi and officeJs attachment arrays
     */
    async prepareAllAttachments(filenames = this.DEFAULT_ATTACHMENTS) {
        try {
            const [graphApiAttachments, officeJsAttachments] = await Promise.all([
                this.prepareGraphApiAttachments(filenames),
                this.prepareOfficeJsAttachments(filenames, false) // Use URLs for Office.js
            ]);
            
            return {
                graphApi: graphApiAttachments,
                officeJs: officeJsAttachments
            };
        } catch (error) {
            console.error('Error preparing attachments:', error);
            return {
                graphApi: [],
                officeJs: []
            };
        }
    },

    /**
     * Get MIME type from filename extension
     * @param {string} filename
     * @returns {string} MIME type
     */
    getContentTypeFromFilename(filename) {
        const ext = filename.split('.').pop().toLowerCase();
        const types = {
            'pdf': 'application/pdf',
            'step': 'application/octet-stream',
            'stp': 'application/octet-stream',
            'dwg': 'application/acad',
            'dxf': 'application/dxf'
        };
        return types[ext] || 'application/octet-stream';
    },

    /**
     * Test if a file is accessible from the backend
     * Returns detailed diagnostic information
     * @param {string} filename - Name of the file to test
     * @returns {Promise<Object>} Diagnostic result
     */
    async testFileAccessibility(filename) {
        const endpoint = `/files/${encodeURIComponent(filename)}`;
        const url = Config.apiUrl + endpoint;
        
        const result = {
            filename: filename,
            url: url,
            success: false,
            status: null,
            statusText: null,
            contentType: null,
            contentLength: null,
            error: null,
            errorDetails: null
        };
        
        try {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 10000); // 10 second timeout for test
            
            const response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Accept': 'application/octet-stream'
                },
                signal: controller.signal
            });
            
            clearTimeout(timeoutId);
            
            result.status = response.status;
            result.statusText = response.statusText;
            result.contentType = response.headers.get('content-type');
            result.contentLength = response.headers.get('content-length');
            
            if (response.ok) {
                const blob = await response.blob();
                result.success = true;
                result.blobSize = blob.size;
            } else {
                // Try to get error message from response
                try {
                    const errorText = await response.text();
                    result.error = `HTTP ${response.status}: ${response.statusText}`;
                    result.errorDetails = errorText.substring(0, 200);
                } catch (e) {
                    result.error = `HTTP ${response.status}: ${response.statusText}`;
                }
            }
        } catch (error) {
            if (error.name === 'AbortError') {
                result.error = 'Request timed out after 10 seconds';
            } else {
                result.error = error.message;
                result.errorDetails = error.toString();
            }
        }
        
        return result;
    },

    /**
     * Test all STEP files and return diagnostic report
     * @param {Array<string>} filenames - Array of filenames to test
     * @returns {Promise<Object>} Diagnostic report
     */
    async runStepFileDiagnostics(filenames) {
        const stepFiles = filenames.filter(f => this.isStepFile(f));
        const results = [];
        
        for (const filename of stepFiles) {
            const result = await this.testFileAccessibility(filename);
            results.push(result);
        }
        
        return {
            totalStepFiles: stepFiles.length,
            tested: results.length,
            successful: results.filter(r => r.success).length,
            failed: results.filter(r => !r.success).length,
            results: results
        };
    },

    /**
     * Fetch a file from backend and convert to base64 for Graph API
     * @param {string} filename - Name of the file
     * @param {string} rfqId - Optional RFQ ID
     * @returns {Promise<string>} Base64-encoded content
     */
    async fetchFileFromBackendAsBase64(filename, rfqId = null) {
        try {
            const blob = await ApiClient.fetchFile(filename, rfqId);
            
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onloadend = () => {
                    const dataUrl = reader.result;
                    const parts = dataUrl.split(',');
                    if (parts.length < 2) {
                        reject(new Error(`Invalid data URL format for ${filename}`));
                        return;
                    }
                    resolve(parts[1]); // Return base64 without data URL prefix
                };
                reader.onerror = reject;
                reader.readAsDataURL(blob);
            });
        } catch (error) {
            console.error(`Error fetching file ${filename} from backend:`, error);
            throw error;
        }
    },

    /**
     * Check if a filename is a STEP file
     * @param {string} filename
     * @returns {boolean}
     */
    isStepFile(filename) {
        const ext = filename.split('.').pop().toLowerCase();
        return ext === 'step' || ext === 'stp';
    },

    /**
     * Log attachment summary with STEP file details
     * @param {Array} attachments - Array of attachment objects
     * @param {string} context - Context for the log (e.g., "before upload")
     */
    logAttachmentSummary(attachments, context = '') {
        const pdfs = attachments.filter(a => a.name && a.name.toLowerCase().endsWith('.pdf'));
        const stepFiles = attachments.filter(a => this.isStepFile(a.name || ''));
        const others = attachments.filter(a => {
            const name = (a.name || '').toLowerCase();
            return !name.endsWith('.pdf') && !this.isStepFile(name);
        });

        console.group(`📎 Attachment Summary ${context ? `(${context})` : ''}`);
        console.log(`Total: ${attachments.length} attachment(s)`);
        console.log(`  PDFs: ${pdfs.length}`);
        console.log(`  STEP files: ${stepFiles.length}`);
        if (others.length > 0) {
            console.log(`  Other files: ${others.length}`);
        }
        
        if (stepFiles.length > 0) {
            console.group('🔧 STEP Files:');
            stepFiles.forEach(att => {
                const size = att.contentBytes ? att.contentBytes.length : 0;
                const sizeKB = (size / 1024).toFixed(2);
                console.log(`  ✓ ${att.name} (${sizeKB} KB base64, type: ${att.contentType || 'unknown'})`);
            });
            console.groupEnd();
        }
        
        if (pdfs.length > 0) {
            console.group('📄 PDFs:');
            pdfs.forEach(att => {
                const size = att.contentBytes ? att.contentBytes.length : 0;
                const sizeKB = (size / 1024).toFixed(2);
                console.log(`  ✓ ${att.name} (${sizeKB} KB base64)`);
            });
            console.groupEnd();
        }
        console.groupEnd();
    },

    /**
     * Prepare attachments from API response filenames
     * @param {Array<string>} filenames - Array of filenames from API
     * @param {string} rfqId - Optional RFQ ID for file fetching
     * @returns {Promise<Object>} Object with {attachments: Array, errors: Array}
     */
    async prepareAttachmentsFromApi(filenames, rfqId = null) {
        const attachments = [];
        const errors = [];
        const stepFilesProcessed = [];
        const stepFilesFailed = [];
        
        // #region agent log
        fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'attachments.js:288',message:'Starting prepareAttachmentsFromApi',data:{filenamesCount:filenames.length,filenames:filenames,rfqId:rfqId},timestamp:Date.now(),sessionId:'debug-session',runId:'run2',hypothesisId:'E'})}).catch(()=>{});
        // #endregion
        
        // Identify STEP files upfront
        const stepFilenames = filenames.filter(f => this.isStepFile(f));
        const nonStepFilenames = filenames.filter(f => !this.isStepFile(f));
        
        if (stepFilenames.length > 0) {
            console.log(`🔧 Detected ${stepFilenames.length} STEP file(s) to process:`, stepFilenames);
        }
        
        for (const filename of filenames) {
            const isStep = this.isStepFile(filename);
            
            if (isStep) {
                console.log(`🔧 [STEP FILE] Starting fetch for: ${filename}`);
            }
            
            try {
                const contentBytes = await this.fetchFileFromBackendAsBase64(filename, rfqId);
                const contentType = this.getContentTypeFromFilename(filename);
                
                // STEP-specific validation
                if (isStep) {
                    if (!contentBytes || contentBytes.length === 0) {
                        throw new Error(`STEP file ${filename} has empty contentBytes`);
                    }
                    if (contentType !== 'application/octet-stream') {
                        console.warn(`⚠️ [STEP FILE] Unexpected content type for ${filename}: ${contentType} (expected application/octet-stream)`);
                    }
                    const sizeKB = (contentBytes.length / 1024).toFixed(2);
                    console.log(`🔧 [STEP FILE] Successfully fetched and encoded: ${filename} (${sizeKB} KB base64)`);
                }
                
                // #region agent log
                fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'attachments.js:293',message:'File fetched and encoded',data:{filename:filename,contentBytesLength:contentBytes.length,contentType:contentType,hasContent:!!contentBytes,isStepFile:isStep},timestamp:Date.now(),sessionId:'debug-session',runId:'run2',hypothesisId:'E'})}).catch(()=>{});
                // #endregion
                
                const attachment = {
                    '@odata.type': '#microsoft.graph.fileAttachment',
                    name: filename,
                    contentType: contentType,
                    contentBytes: contentBytes
                };
                
                attachments.push(attachment);
                
                if (isStep) {
                    stepFilesProcessed.push(filename);
                    console.log(`✓ [STEP FILE] Prepared attachment: ${filename}`);
                } else {
                    console.log(`✓ Prepared attachment from API: ${filename}`);
                }
            } catch (error) {
                // Record error details
                const errorInfo = {
                    filename: filename,
                    isStepFile: isStep,
                    errorType: error.name || 'Unknown',
                    errorMessage: error.message || 'Unknown error',
                    errorStack: error.stack ? error.stack.substring(0, 500) : null
                };
                errors.push(errorInfo);
                
                if (isStep) {
                    stepFilesFailed.push({ filename, error: error.message });
                    console.error(`✗ [STEP FILE] CRITICAL: Failed to prepare STEP file ${filename}:`, error);
                    console.error(`  Error type: ${error.name}, Message: ${error.message}`);
                } else {
                    console.error(`✗ Failed to prepare attachment ${filename}:`, error);
                    console.error(`  Error type: ${error.name}, Message: ${error.message}`);
                }
                
                if (error.stack) {
                    console.error(`  Stack: ${error.stack.substring(0, 200)}...`);
                }
                
                // #region agent log
                fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'attachments.js:305',message:'Failed to prepare attachment',data:{filename:filename,errorMessage:error.message,errorName:error.name,isStepFile:isStep},timestamp:Date.now(),sessionId:'debug-session',runId:'run2',hypothesisId:'E'})}).catch(()=>{});
                // #endregion
                
                // Continue with other attachments - don't fail entire operation
            }
        }
        
        // STEP file summary
        if (stepFilenames.length > 0) {
            console.group('🔧 STEP File Processing Summary');
            console.log(`Total STEP files in request: ${stepFilenames.length}`);
            console.log(`Successfully processed: ${stepFilesProcessed.length}`);
            console.log(`Failed: ${stepFilesFailed.length}`);
            if (stepFilesProcessed.length > 0) {
                console.log(`✓ Successful STEP files:`, stepFilesProcessed);
            }
            if (stepFilesFailed.length > 0) {
                console.error(`✗ Failed STEP files:`, stepFilesFailed.map(f => `${f.filename} (${f.error})`));
            }
            console.groupEnd();
        }
        
        // Verify STEP files are in final array
        const stepFilesInAttachments = attachments.filter(a => this.isStepFile(a.name));
        if (stepFilenames.length > 0 && stepFilesInAttachments.length !== stepFilesProcessed.length) {
            console.warn(`⚠️ WARNING: STEP file count mismatch! Expected ${stepFilesProcessed.length} in attachments array, found ${stepFilesInAttachments.length}`);
        }
        
        // #region agent log
        fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'attachments.js:310',message:'prepareAttachmentsFromApi completed',data:{totalAttachments:attachments.length,attachmentNames:attachments.map(a=>a.name),stepFilesCount:stepFilesInAttachments.length,stepFilesProcessed:stepFilesProcessed.length,stepFilesFailed:stepFilesFailed.length},timestamp:Date.now(),sessionId:'debug-session',runId:'run2',hypothesisId:'E'})}).catch(()=>{});
        // #endregion
        
        // Log summary
        this.logAttachmentSummary(attachments, 'after API fetch');
        
        // Return both attachments and errors for UI feedback
        return {
            attachments: attachments,
            errors: errors,
            stepFilesProcessed: stepFilesProcessed,
            stepFilesFailed: stepFilesFailed.map(f => f.filename)
        };
    }
};
