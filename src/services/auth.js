/**
 * Authentication Service using MSAL.js
 * Handles Microsoft Graph API authentication
 */
const AuthService = {
    // MSAL configuration
    msalConfig: {
        auth: {
            clientId: '6492ece4-5f90-4781-bc60-3cb5fc5adb10',
            authority: 'https://login.microsoftonline.com/common',
            // Automatically detect redirect URI based on current location
            redirectUri: window.location.origin + '/src/taskpane/taskpane.html',
        },
        cache: {
            cacheLocation: 'localStorage',
            storeAuthStateInCookie: false,
        }
    },

    // Scopes needed for Graph API
    scopes: [
        'User.Read',
        'Mail.ReadWrite',
        'Mail.Send',
        'MailboxFolder.ReadWrite',
        'MailboxSettings.ReadWrite'  // Required for creating/managing email categories
    ],

    // MSAL instance
    msalInstance: null,

    // Current account
    currentAccount: null,

    /**
     * Initialize MSAL
     */
    async initialize() {
        if (typeof msal === 'undefined') {
            console.error('MSAL library not loaded');
            return false;
        }

        try {
            this.msalInstance = new msal.PublicClientApplication(this.msalConfig);
            
            // Handle redirect response
            const response = await this.msalInstance.handleRedirectPromise();
            if (response) {
                this.currentAccount = response.account;
                console.log('Logged in via redirect:', this.currentAccount.username);
            } else {
                // Check for existing accounts
                const accounts = this.msalInstance.getAllAccounts();
                if (accounts.length > 0) {
                    this.currentAccount = accounts[0];
                    console.log('Found existing account:', this.currentAccount.username);
                }
            }

            return true;
        } catch (error) {
            console.error('MSAL initialization error:', error);
            return false;
        }
    },

    /**
     * Check if user is signed in
     */
    isSignedIn() {
        return this.currentAccount !== null;
    },

    /**
     * Get current user info
     */
    getUser() {
        if (!this.currentAccount) return null;
        return {
            name: this.currentAccount.name,
            email: this.currentAccount.username,
            id: this.currentAccount.localAccountId
        };
    },

    /**
     * Sign in with popup
     */
    async signIn() {
        if (!this.msalInstance) {
            throw new Error('MSAL not initialized');
        }

        try {
            const response = await this.msalInstance.loginPopup({
                scopes: this.scopes,
                prompt: 'select_account'
            });
            
            this.currentAccount = response.account;
            console.log('Signed in:', this.currentAccount.username);
            
            return this.getUser();
        } catch (error) {
            console.error('Sign in error:', error);
            throw error;
        }
    },

    /**
     * Sign out
     */
    async signOut() {
        if (!this.msalInstance || !this.currentAccount) {
            return;
        }

        try {
            await this.msalInstance.logoutPopup({
                account: this.currentAccount,
                postLogoutRedirectUri: this.msalConfig.auth.redirectUri
            });
            this.currentAccount = null;
            console.log('Signed out');
        } catch (error) {
            console.error('Sign out error:', error);
            throw error;
        }
    },

    /**
     * Get access token for Graph API
     */
    async getAccessToken() {
        if (!this.msalInstance) {
            throw new Error('MSAL not initialized');
        }

        if (!this.currentAccount) {
            throw new Error('No user signed in');
        }

        const tokenRequest = {
            scopes: this.scopes,
            account: this.currentAccount
        };

        try {
            // Try silent token acquisition first
            const response = await this.msalInstance.acquireTokenSilent(tokenRequest);
            return response.accessToken;
        } catch (error) {
            // If silent fails, try popup
            if (error instanceof msal.InteractionRequiredAuthError) {
                console.log('Silent token acquisition failed, using popup');
                const response = await this.msalInstance.acquireTokenPopup(tokenRequest);
                return response.accessToken;
            }
            throw error;
        }
    },

    /**
     * Make authenticated Graph API request
     */
    async graphRequest(endpoint, options = {}) {
        const token = await this.getAccessToken();
        
        const url = endpoint.startsWith('https://') 
            ? endpoint 
            : `https://graph.microsoft.com/v1.0${endpoint}`;

        const response = await fetch(url, {
            ...options,
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json',
                ...options.headers
            }
        });

        if (!response.ok) {
            const error = await response.json().catch(() => ({}));
            throw new Error(`Graph API error: ${error.error?.message || response.statusText}`);
        }

        // Handle 204 No Content and 202 Accepted (sendMail returns 202)
        if (response.status === 204 || response.status === 202) {
            return null;
        }

        // Check if response has content before parsing JSON
        const contentType = response.headers.get('content-type');
        if (contentType && contentType.includes('application/json')) {
            const text = await response.text();
            if (text) {
                return JSON.parse(text);
            }
        }
        
        return null;
    }
};