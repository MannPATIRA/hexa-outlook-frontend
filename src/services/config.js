/**
 * Configuration settings for the Procurement Workflow Add-in
 */
const Config = {
    // Backend API URL - ALWAYS use deployed backend URL
    // NEVER use localhost - always use production URL
    API_BASE_URL: (() => {
        // Always use deployed backend - never localhost
        return 'https://hexa-outlook-backend.onrender.com';
    })(),
    API_PREFIX: '/api',

    // Get the full API URL
    get apiUrl() {
        return this.API_BASE_URL + this.API_PREFIX;
    },

    // Engineering team email for forwarding clarifications
    ENGINEERING_EMAIL: 'engineering@company.com',

    // Folder names for email organization
    FOLDERS: {
        SENT_RFQS: 'Sent RFQs',
        QUOTES: 'Quotes',
        CLARIFICATION_REQUESTS: 'Clarification Requests',
        AWAITING_CLARIFICATION: 'Awaiting Clarification Response',
        AWAITING_ENGINEER: 'Awaiting Engineer Response',
        ENGINEER_RESPONSE: 'Engineer Response'
    },

    // Folder-based category definitions for email tagging
    // Each folder maps to a category name and color preset
    // Colors: Preset0=Red, Preset1=Orange, Preset3=Yellow, Preset4=Green, Preset5=Teal, Preset7=Blue, Preset8=Purple
    FOLDER_CATEGORIES: {
        'Sent RFQs': { name: 'Sent RFQ', color: 'Preset7' },                          // Blue
        'Quotes': { name: 'Quote', color: 'Preset4' },                                 // Green
        'Clarification Requests': { name: 'Clarification', color: 'Preset3' },         // Yellow
        'Awaiting Clarification Response': { name: 'Awaiting Response', color: 'Preset1' },  // Orange
        'Awaiting Engineer Response': { name: 'Awaiting Engineer', color: 'Preset8' },       // Purple
        'Engineer Response': { name: 'Engineer Response', color: 'Preset5' }           // Teal
    },

    // Local storage keys
    STORAGE_KEYS: {
        API_URL: 'procurement_api_url',
        ENGINEERING_EMAIL: 'procurement_engineering_email',
        AUTO_CLASSIFY: 'procurement_auto_classify',
        AUTO_CREATE_FOLDERS: 'procurement_auto_create_folders',
        CACHED_PRS: 'procurement_cached_prs',
        CACHED_SUPPLIERS: 'procurement_cached_suppliers'
    },

    // Request timeout in milliseconds
    REQUEST_TIMEOUT: 30000,

    // OpenAI API Configuration
    // Priority: 1. Environment variable (from Vercel/build), 2. localStorage, 3. Empty string
    get OPENAI_API_KEY() {
        // First try environment variable (injected at build time for Vercel)
        // This is set from process.env.OPENAI_API_KEY during build
        if (typeof window !== 'undefined' && window.OPENAI_API_KEY_ENV) {
            return window.OPENAI_API_KEY_ENV;
        }
        
        // Fallback to localStorage (for user configuration or local dev)
        try {
            const storedKey = localStorage.getItem('procurement_openai_api_key');
            if (storedKey) return storedKey;
        } catch (e) {
            // localStorage might not be available in some contexts
            console.warn('localStorage not available:', e);
        }
        
        // Return empty string if not configured (will cause API calls to fail)
        return '';
    },
    OPENAI_API_BASE_URL: 'https://api.openai.com/v1',

    // Load settings from local storage
    loadSettings() {
        try {
            const apiUrl = localStorage.getItem(this.STORAGE_KEYS.API_URL);
            // Only use stored URL if it's not localhost - always prefer production
            if (apiUrl && !apiUrl.includes('localhost') && !apiUrl.includes('127.0.0.1')) {
                this.API_BASE_URL = apiUrl;
            }
            // If stored URL is localhost, ignore it and use production default
            else if (apiUrl && (apiUrl.includes('localhost') || apiUrl.includes('127.0.0.1'))) {
                console.warn('Ignoring localhost URL from localStorage, using production URL');
                // Clear the localhost URL from storage
                localStorage.removeItem(this.STORAGE_KEYS.API_URL);
            }

            const engineeringEmail = localStorage.getItem(this.STORAGE_KEYS.ENGINEERING_EMAIL);
            if (engineeringEmail) this.ENGINEERING_EMAIL = engineeringEmail;
        } catch (e) {
            console.log('Could not load settings from localStorage:', e);
        }
    },

    // Save settings to local storage
    saveSettings(settings) {
        try {
            if (settings.apiUrl) {
                // Prevent saving localhost URLs - always use production
                if (settings.apiUrl.includes('localhost') || settings.apiUrl.includes('127.0.0.1')) {
                    console.warn('Cannot save localhost URL. Using production URL instead.');
                    // Don't save localhost, but continue to save other settings
                } else {
                    this.API_BASE_URL = settings.apiUrl;
                    localStorage.setItem(this.STORAGE_KEYS.API_URL, settings.apiUrl);
                }
            }
            if (settings.engineeringEmail) {
                this.ENGINEERING_EMAIL = settings.engineeringEmail;
                localStorage.setItem(this.STORAGE_KEYS.ENGINEERING_EMAIL, settings.engineeringEmail);
            }
            if (typeof settings.autoClassify !== 'undefined') {
                localStorage.setItem(this.STORAGE_KEYS.AUTO_CLASSIFY, settings.autoClassify);
            }
            if (typeof settings.autoCreateFolders !== 'undefined') {
                localStorage.setItem(this.STORAGE_KEYS.AUTO_CREATE_FOLDERS, settings.autoCreateFolders);
            }
        } catch (e) {
            console.log('Could not save settings to localStorage:', e);
        }
    },

    // Get setting value
    getSetting(key, defaultValue = null) {
        try {
            const value = localStorage.getItem(key);
            if (value === null) return defaultValue;
            if (value === 'true') return true;
            if (value === 'false') return false;
            return value;
        } catch (e) {
            return defaultValue;
        }
    },

    // Set a single setting value
    setSetting(key, value) {
        try {
            localStorage.setItem(key, value);
        } catch (e) {
            console.log('Could not save setting to localStorage:', e);
        }
    }
};

// Load settings on initialization
Config.loadSettings();
