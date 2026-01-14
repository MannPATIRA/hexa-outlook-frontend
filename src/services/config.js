/**
 * Configuration settings for the Procurement Workflow Add-in
 */
const Config = {
    // Backend API URL - change this for different environments
    // For production, set this to your deployed backend URL (e.g., https://your-backend.railway.app)
    // For local development, use http://localhost:8000
    // This will be overridden by localStorage if user configures it in settings
    API_BASE_URL: (() => {
        // Check if we're in production (not localhost)
        if (window.location.hostname !== 'localhost' && window.location.hostname !== '127.0.0.1') {
            // Production: You need to set this to your actual backend URL
            // TODO: Replace with your deployed backend URL
            return 'https://your-backend.railway.app'; // CHANGE THIS
        }
        // Development: use localhost
        return 'http://localhost:8000';
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
            if (apiUrl) this.API_BASE_URL = apiUrl;

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
                this.API_BASE_URL = settings.apiUrl;
                localStorage.setItem(this.STORAGE_KEYS.API_URL, settings.apiUrl);
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
