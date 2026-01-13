/**
 * Utility Helper Functions
 */
const Helpers = {
    /**
     * Format a date for display
     */
    formatDate(date, includeTime = false) {
        if (!date) return 'N/A';
        
        const d = new Date(date);
        const options = {
            year: 'numeric',
            month: 'short',
            day: 'numeric'
        };
        
        if (includeTime) {
            options.hour = '2-digit';
            options.minute = '2-digit';
        }
        
        return d.toLocaleDateString('en-US', options);
    },

    /**
     * Format currency
     */
    formatCurrency(amount, currency = 'USD') {
        if (amount === null || amount === undefined) return 'N/A';
        
        return new Intl.NumberFormat('en-US', {
            style: 'currency',
            currency: currency
        }).format(amount);
    },

    /**
     * Truncate text with ellipsis
     */
    truncate(text, maxLength = 50) {
        if (!text) return '';
        if (text.length <= maxLength) return text;
        return text.substring(0, maxLength - 3) + '...';
    },

    /**
     * Escape HTML to prevent XSS
     */
    escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    },

    /**
     * Parse HTML to plain text
     */
    htmlToText(html) {
        if (!html) return '';
        const div = document.createElement('div');
        div.innerHTML = html;
        return div.textContent || div.innerText || '';
    },

    /**
     * Strip HTML tags from text (alias for htmlToText)
     */
    stripHtml(html) {
        return this.htmlToText(html);
    },

    /**
     * Debounce function calls
     */
    debounce(func, wait = 300) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    },

    /**
     * Generate a unique ID
     */
    generateId(prefix = 'id') {
        return `${prefix}_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    },

    /**
     * Deep clone an object
     */
    deepClone(obj) {
        return JSON.parse(JSON.stringify(obj));
    },

    /**
     * Check if object is empty
     */
    isEmpty(obj) {
        if (!obj) return true;
        if (Array.isArray(obj)) return obj.length === 0;
        if (typeof obj === 'object') return Object.keys(obj).length === 0;
        return false;
    },

    /**
     * Get match score class for styling
     */
    getMatchScoreClass(score) {
        if (score >= 8) return 'high';
        if (score >= 5) return 'medium';
        return 'low';
    },

    /**
     * Sort array of objects by property
     */
    sortBy(array, property, descending = false) {
        return [...array].sort((a, b) => {
            const aVal = a[property];
            const bVal = b[property];
            
            if (aVal < bVal) return descending ? 1 : -1;
            if (aVal > bVal) return descending ? -1 : 1;
            return 0;
        });
    },

    /**
     * Filter array based on search term
     */
    filterBySearch(array, searchTerm, ...properties) {
        if (!searchTerm) return array;
        
        const term = searchTerm.toLowerCase();
        return array.filter(item => {
            return properties.some(prop => {
                const value = item[prop];
                if (!value) return false;
                return String(value).toLowerCase().includes(term);
            });
        });
    },

    /**
     * Show/hide element
     */
    showElement(element) {
        if (element) element.classList.remove('hidden');
    },

    /**
     * Hide element
     */
    hideElement(element) {
        if (element) element.classList.add('hidden');
    },

    /**
     * Toggle element visibility
     */
    toggleElement(element, show) {
        if (!element) return;
        if (show === undefined) {
            element.classList.toggle('hidden');
        } else if (show) {
            this.showElement(element);
        } else {
            this.hideElement(element);
        }
    },

    /**
     * Enable/disable workflow step
     */
    enableStep(stepElement) {
        if (stepElement) stepElement.classList.remove('disabled');
    },

    /**
     * Disable workflow step
     */
    disableStep(stepElement) {
        if (stepElement) stepElement.classList.add('disabled');
    },

    /**
     * Set button loading state
     */
    setButtonLoading(button, loading, originalText = null) {
        if (!button) return;
        
        if (loading) {
            button.dataset.originalText = button.textContent;
            button.disabled = true;
            button.innerHTML = '<span class="spinner-small"></span> Loading...';
        } else {
            button.disabled = false;
            button.textContent = originalText || button.dataset.originalText || 'Submit';
        }
    },

    /**
     * Display notification banner
     */
    showNotification(message, type = 'success') {
        const banner = document.getElementById(`${type}-banner`);
        const messageEl = document.getElementById(`${type}-message`);
        
        if (banner && messageEl) {
            messageEl.textContent = message;
            this.showElement(banner);
            
            // Auto-hide after 5 seconds
            setTimeout(() => {
                this.hideElement(banner);
            }, 5000);
        }
    },

    /**
     * Show error notification
     */
    showError(message) {
        this.showNotification(message, 'error');
    },

    /**
     * Show success notification
     */
    showSuccess(message) {
        this.showNotification(message, 'success');
    },

    /**
     * Show loading overlay
     */
    showLoading(message = 'Loading...') {
        const overlay = document.getElementById('loading-overlay');
        const messageEl = document.getElementById('loading-message');
        
        if (overlay) {
            if (messageEl) messageEl.textContent = message;
            this.showElement(overlay);
        }
    },

    /**
     * Hide loading overlay
     */
    hideLoading() {
        const overlay = document.getElementById('loading-overlay');
        if (overlay) this.hideElement(overlay);
    },

    /**
     * Create element with attributes and content
     */
    createElement(tag, attributes = {}, content = '') {
        const element = document.createElement(tag);
        
        for (const [key, value] of Object.entries(attributes)) {
            if (key === 'className') {
                element.className = value;
            } else if (key === 'dataset') {
                for (const [dataKey, dataValue] of Object.entries(value)) {
                    element.dataset[dataKey] = dataValue;
                }
            } else if (key.startsWith('on')) {
                element.addEventListener(key.substring(2).toLowerCase(), value);
            } else {
                element.setAttribute(key, value);
            }
        }
        
        if (typeof content === 'string') {
            element.innerHTML = content;
        } else if (content instanceof Element) {
            element.appendChild(content);
        } else if (Array.isArray(content)) {
            content.forEach(child => {
                if (child instanceof Element) {
                    element.appendChild(child);
                } else {
                    element.innerHTML += child;
                }
            });
        }
        
        return element;
    },

    /**
     * Clear all children of an element
     */
    clearChildren(element) {
        if (element) {
            while (element.firstChild) {
                element.removeChild(element.firstChild);
            }
        }
    },

    /**
     * Validate email address
     */
    isValidEmail(email) {
        const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        return re.test(email);
    },

    /**
     * Extract material code from PR data
     */
    extractMaterialCode(pr) {
        // Try different property names
        return pr.material || pr.material_code || pr.materialCode || 
               pr.pr_id || 'UNKNOWN';
    },

    /**
     * Get classification display name
     */
    getClassificationDisplayName(classification) {
        const names = {
            'quote': 'Quote',
            'clarification_request': 'Clarification Request',
            'engineer_response': 'Engineer Response',
            'procurement': 'Procurement',
            'engineering': 'Engineering'
        };
        return names[classification] || classification;
    },

    /**
     * Format confidence as percentage
     */
    formatConfidence(confidence) {
        if (confidence === null || confidence === undefined) return 'N/A';
        return `${Math.round(confidence * 100)}% confidence`;
    }
};
