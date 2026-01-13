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
    },

    /**
     * Parse clarification questions from email body
     * Extracts individual questions from structured text (numbered sections, bullet points, etc.)
     * @param {string} emailBody - The email body text (HTML or plain text)
     * @returns {Array<Object>} Array of question objects with {category, question} structure
     */
    parseClarificationQuestions(emailBody) {
        if (!emailBody) return [];
        
        // Convert HTML to plain text if needed
        let text = this.stripHtml(emailBody);
        
        // Normalize line breaks
        text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        
        const questions = [];
        
        // Pattern 1: Numbered sections with bullet points (e.g., "1. Tolerance Requirements\n- Question 1\n- Question 2")
        const numberedSectionPattern = /(\d+)\.\s*([^\n]+(?:\n(?!\d+\.)[^\n]*)*)/g;
        let match;
        
        while ((match = numberedSectionPattern.exec(text)) !== null) {
            const sectionNumber = match[1];
            const sectionContent = match[2].trim();
            
            // Extract category/title (first line)
            const lines = sectionContent.split('\n').map(l => l.trim()).filter(l => l);
            if (lines.length === 0) continue;
            
            const category = lines[0];
            
            // Extract individual questions (lines starting with -, •, *, or numbered sub-items)
            const questionPattern = /^[-•*]\s*(.+)$|^(\d+[\.\)])\s*(.+)$/;
            
            for (let i = 1; i < lines.length; i++) {
                const line = lines[i];
                const questionMatch = line.match(questionPattern);
                
                if (questionMatch) {
                    // Found a bullet point or sub-question
                    const questionText = questionMatch[1] || questionMatch[3] || line;
                    questions.push({
                        category: category,
                        question: questionText.trim(),
                        sectionNumber: sectionNumber
                    });
                } else if (line.length > 10 && !line.match(/^[A-Z\s]+$/)) {
                    // If line doesn't look like a header and is substantial, treat as a question
                    questions.push({
                        category: category,
                        question: line.trim(),
                        sectionNumber: sectionNumber
                    });
                }
            }
            
            // If no sub-questions found but category exists, add category as a question
            if (lines.length === 1 && category.length > 5) {
                questions.push({
                    category: category,
                    question: category,
                    sectionNumber: sectionNumber
                });
            }
        }
        
        // Pattern 2: Simple bullet points without numbered sections
        if (questions.length === 0) {
            const bulletPattern = /^[-•*]\s*(.+)$/gm;
            let bulletMatch;
            
            while ((bulletMatch = bulletPattern.exec(text)) !== null) {
                const questionText = bulletMatch[1].trim();
                if (questionText.length > 5) { // Filter out very short items
                    questions.push({
                        category: 'General Questions',
                        question: questionText,
                        sectionNumber: null
                    });
                }
            }
        }
        
        // Pattern 3: Questions separated by line breaks (fallback)
        if (questions.length === 0) {
            // Look for lines that end with "?" or are substantial questions
            const questionLines = text.split('\n')
                .map(l => l.trim())
                .filter(l => l.length > 10 && (l.includes('?') || l.match(/^(what|how|why|when|where|can|could|would|is|are|do|does)/i)));
            
            questionLines.forEach((line, index) => {
                questions.push({
                    category: 'Questions',
                    question: line,
                    sectionNumber: (index + 1).toString()
                });
            });
        }
        
        // Pattern 4: If still no questions found, split by common separators
        if (questions.length === 0) {
            // Try splitting by double newlines or common question markers
            const sections = text.split(/\n\n+/).filter(s => s.trim().length > 10);
            
            sections.forEach((section, index) => {
                const cleanSection = section.trim().replace(/\n/g, ' ').substring(0, 200);
                if (cleanSection.length > 10) {
                    questions.push({
                        category: 'Question',
                        question: cleanSection,
                        sectionNumber: (index + 1).toString()
                    });
                }
            });
        }
        
        // Remove duplicates and clean up
        const uniqueQuestions = [];
        const seen = new Set();
        
        questions.forEach(q => {
            const key = q.question.toLowerCase().trim();
            if (!seen.has(key) && q.question.length > 5) {
                seen.add(key);
                uniqueQuestions.push(q);
            }
        });
        
        return uniqueQuestions;
    },

    /**
     * Wrap a promise with a timeout
     * @param {Promise} promise - The promise to wrap
     * @param {number} ms - Timeout in milliseconds
     * @param {string} errorMessage - Error message if timeout occurs
     * @returns {Promise} Promise that rejects if timeout is exceeded
     */
    withTimeout(promise, ms, errorMessage = 'Operation timed out') {
        return Promise.race([
            promise,
            new Promise((_, reject) => 
                setTimeout(() => reject(new Error(errorMessage)), ms)
            )
        ]);
    }
};
