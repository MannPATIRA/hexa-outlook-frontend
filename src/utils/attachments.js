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
    }
};
