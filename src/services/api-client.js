/**
 * API Client for Procurement Backend
 * Handles all HTTP requests to the FastAPI backend
 */
const ApiClient = {
    /**
     * Make a fetch request with error handling
     */
    async request(endpoint, options = {}) {
        const url = Config.apiUrl + endpoint;
        
        const defaultOptions = {
            headers: {
                'Content-Type': 'application/json',
            },
        };

        const fetchOptions = { ...defaultOptions, ...options };
        
        try {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), Config.REQUEST_TIMEOUT);
            
            const response = await fetch(url, {
                ...fetchOptions,
                signal: controller.signal
            });
            
            clearTimeout(timeoutId);

            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                throw new ApiError(
                    errorData.detail || `HTTP ${response.status}: ${response.statusText}`,
                    response.status,
                    errorData
                );
            }

            return await response.json();
        } catch (error) {
            if (error.name === 'AbortError') {
                throw new ApiError('Request timed out', 408);
            }
            if (error instanceof ApiError) {
                throw error;
            }
            throw new ApiError(
                `Network error: ${error.message}. Is the backend server running at ${Config.API_BASE_URL}?`,
                0
            );
        }
    },

    /**
     * GET request helper
     */
    async get(endpoint) {
        return this.request(endpoint, { method: 'GET' });
    },

    /**
     * POST request helper
     */
    async post(endpoint, data) {
        return this.request(endpoint, {
            method: 'POST',
            body: JSON.stringify(data)
        });
    },

    // ==================== PR ENDPOINTS ====================

    /**
     * Get all open Purchase Requisitions
     */
    async getOpenPRs() {
        const response = await this.get('/prs/open');
        return response.prs || [];
    },

    // ==================== SUPPLIER ENDPOINTS ====================

    /**
     * Search for suppliers matching a PR
     */
    async searchSuppliers(prId, material = null, specs = null) {
        const payload = { pr_id: prId };
        if (material) payload.material = material;
        if (specs) payload.specs = specs;
        
        const response = await this.post('/suppliers/search', payload);
        return response.suppliers || [];
    },

    // ==================== RFQ ENDPOINTS ====================

    /**
     * Generate RFQs for selected suppliers
     */
    async generateRFQs(prId, supplierIds) {
        const response = await this.post('/rfqs/generate', {
            pr_id: prId,
            supplier_ids: supplierIds
        });
        // #region agent log
        fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'api-client.js:106',message:'API response received',data:{hasResponse:!!response,hasRfqs:!!response.rfqs,rfqsLength:response.rfqs?.length,firstRfqAttachments:response.rfqs?.[0]?.attachments,firstRfqAttachmentsType:typeof response.rfqs?.[0]?.attachments,firstRfqAttachmentsIsArray:Array.isArray(response.rfqs?.[0]?.attachments),firstRfqBody:response.rfqs?.[0]?.body,fullResponse:JSON.stringify(response).substring(0,500)},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'A'})}).catch(()=>{});
        // #endregion
        return response.rfqs || [];
    },

    /**
     * Finalize an RFQ with edited content
     */
    async finalizeRFQ(rfqId, finalSubject, finalBody, status = 'ready_to_send') {
        return this.post('/rfqs/finalize', {
            rfq_id: rfqId,
            final_subject: finalSubject,
            final_body: finalBody,
            status: status
        });
    },

    // ==================== EMAIL ENDPOINTS ====================

    /**
     * Classify an incoming email
     */
    async classifyEmail(emailChain, mostRecentReply, rfqId = null, supplierId = null) {
        const payload = {
            email_chain: emailChain,
            most_recent_reply: mostRecentReply
        };
        if (rfqId) payload.rfq_id = rfqId;
        if (supplierId) payload.supplier_id = supplierId;
        
        return this.post('/emails/classify', payload);
    },

    /**
     * Process a classified email
     */
    async processEmail(emailId, classification) {
        return this.post('/emails/process', {
            email_id: emailId,
            classification: classification
        });
    },

    /**
     * Get suggested response for a clarification
     */
    async suggestResponse(clarificationId, emailId, question) {
        return this.post('/emails/suggest-response', {
            clarification_id: clarificationId,
            email_id: emailId,
            question: question
        });
    },

    /**
     * Forward clarification to engineering
     */
    async forwardToEngineering(emailId, clarificationId) {
        return this.post('/emails/forward-to-engineering', {
            email_id: emailId,
            clarification_id: clarificationId
        });
    },

    /**
     * Process engineer's response
     */
    async processEngineerResponse(emailId, engineerResponse) {
        return this.post('/emails/process-engineer-response', {
            email_id: emailId,
            engineer_response: engineerResponse
        });
    },

    /**
     * Extract quote data from email
     */
    async extractQuote(emailId, rfqId, supplierId, emailBody) {
        return this.post('/emails/extract-quote', {
            email_id: emailId,
            rfq_id: rfqId,
            supplier_id: supplierId,
            email_body: emailBody
        });
    },

    // ==================== QUOTE ENDPOINTS ====================

    /**
     * Get all quotes for an RFQ
     */
    async getQuotes(rfqId) {
        const response = await this.get(`/quotes/${rfqId}`);
        return response.quotes || [];
    },

    // ==================== DEMO ENDPOINTS ====================

    /**
     * Schedule an automatic reply for demo/testing purposes
     * The backend will send a simulated supplier reply after the specified delay
     * @param {Object} options - Reply options
     * @param {string} options.toEmail - User's email address (where to send the reply)
     * @param {string} options.subject - Original RFQ subject
     * @param {string} options.internetMessageId - Internet Message ID of the sent RFQ (for threading)
     * @param {string} options.material - Material name for generating realistic reply content
     * @param {string} options.replyType - Type of reply: "quote", "clarification_procurement", "clarification_engineering", or "random"
     * @param {number} options.delaySeconds - How long to wait before sending (default: 30)
     * @param {number} options.quantity - Quantity for quote calculations (default: 100)
     */
    async scheduleAutoReply(options) {
        return this.post('/demo/schedule-reply', {
            to_email: options.toEmail,
            original_subject: options.subject,
            original_message_id: options.internetMessageId,
            material: options.material,
            reply_type: options.replyType || 'random',
            delay_seconds: options.delaySeconds || 30,
            quantity: options.quantity || 100
        });
    },

    // ==================== FILE ENDPOINTS ====================

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
     * Fetch a file from the backend file server
     * @param {string} filename - Name of the file to fetch
     * @param {string} rfqId - Optional RFQ ID for context
     * @returns {Promise<Blob>} File blob
     */
    async fetchFile(filename, rfqId = null) {
        const isStep = this.isStepFile(filename);
        const endpoint = rfqId 
            ? `/files/rfq/${rfqId}/${encodeURIComponent(filename)}`
            : `/files/${encodeURIComponent(filename)}`;
        
        const url = Config.apiUrl + endpoint;
        
        if (isStep) {
            console.log(`🔧 [STEP FILE] Fetching from backend: ${filename}`);
            console.log(`🔧 [STEP FILE] URL: ${url}`);
        }
        
        // #region agent log
        fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'api-client.js:234',message:'Attempting to fetch file',data:{filename:filename,rfqId:rfqId,endpoint:endpoint,fullUrl:url,isStepFile:isStep},timestamp:Date.now(),sessionId:'debug-session',runId:'run2',hypothesisId:'E'})}).catch(()=>{});
        // #endregion
        
        try {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), Config.REQUEST_TIMEOUT);
            
            const response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Accept': 'application/octet-stream'
                },
                signal: controller.signal
            });
            
            clearTimeout(timeoutId);
            
            const contentType = response.headers.get('content-type');
            const contentLength = response.headers.get('content-length');
            
            if (isStep) {
                console.log(`🔧 [STEP FILE] Response received for ${filename}:`);
                console.log(`  Status: ${response.status} ${response.statusText}`);
                console.log(`  Content-Type: ${contentType || 'not set'}`);
                console.log(`  Content-Length: ${contentLength ? (parseInt(contentLength) / 1024).toFixed(2) + ' KB' : 'not set'}`);
            }
            
            // #region agent log
            fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'api-client.js:248',message:'File fetch response received',data:{filename:filename,status:response.status,statusText:response.statusText,ok:response.ok,contentType:contentType,contentLength:contentLength,isStepFile:isStep},timestamp:Date.now(),sessionId:'debug-session',runId:'run2',hypothesisId:'E'})}).catch(()=>{});
            // #endregion
            
            if (!response.ok) {
                if (isStep) {
                    console.error(`🔧 [STEP FILE] CRITICAL: Fetch failed for ${filename}: ${response.status} ${response.statusText}`);
                }
                // #region agent log
                fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'api-client.js:252',message:'File fetch failed',data:{filename:filename,status:response.status,statusText:response.statusText,url:url,isStepFile:isStep},timestamp:Date.now(),sessionId:'debug-session',runId:'run2',hypothesisId:'E'})}).catch(()=>{});
                // #endregion
                throw new ApiError(`Failed to fetch file ${filename}: ${response.statusText}`, response.status);
            }
            
            const blob = await response.blob();
            
            if (isStep) {
                const sizeKB = (blob.size / 1024).toFixed(2);
                console.log(`🔧 [STEP FILE] Blob received for ${filename}: ${sizeKB} KB (type: ${blob.type || 'unknown'})`);
            }
            
            // #region agent log
            fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'api-client.js:260',message:'File blob received',data:{filename:filename,blobSize:blob.size,blobType:blob.type,isStepFile:isStep},timestamp:Date.now(),sessionId:'debug-session',runId:'run2',hypothesisId:'E'})}).catch(()=>{});
            // #endregion
            
            return blob;
        } catch (error) {
            if (isStep) {
                console.error(`🔧 [STEP FILE] CRITICAL: Error fetching ${filename}:`, error);
                console.error(`  Error type: ${error.name}`);
                console.error(`  Error message: ${error.message}`);
            }
            // #region agent log
            fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'api-client.js:264',message:'File fetch error',data:{filename:filename,errorName:error.name,errorMessage:error.message,errorStack:error.stack,url:url,isStepFile:isStep},timestamp:Date.now(),sessionId:'debug-session',runId:'run2',hypothesisId:'E'})}).catch(()=>{});
            // #endregion
            if (error.name === 'AbortError') {
                throw new ApiError(`Request timed out while fetching file ${filename}`, 408);
            }
            if (error instanceof ApiError) {
                throw error;
            }
            throw new ApiError(
                `Network error fetching file ${filename}: ${error.message}`,
                0
            );
        }
    }
};

/**
 * Custom API Error class
 */
class ApiError extends Error {
    constructor(message, statusCode, data = null) {
        super(message);
        this.name = 'ApiError';
        this.statusCode = statusCode;
        this.data = data;
    }
}

// Make ApiError available globally
window.ApiError = ApiError;
