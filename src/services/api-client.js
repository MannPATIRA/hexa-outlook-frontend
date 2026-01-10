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
