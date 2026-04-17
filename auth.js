// BlackIP360 Présences — Authentification MSAL.js (redirect-only)
class BlackIPAuth {
  constructor() {
    this.msal    = null;
    this.account = null;
    this.initError = null;
  }

  async init() {
    // L'URI de redirect doit correspondre EXACTEMENT à ce qui est enregistré dans Azure AD
    const redirectUri = window.location.origin + window.location.pathname;

    const msalCfg = {
      auth: {
        clientId:                  CONFIG.CLIENT_ID,
        authority:                 `https://login.microsoftonline.com/${CONFIG.TENANT_ID}`,
        redirectUri:               redirectUri,
        postLogoutRedirectUri:     redirectUri,
        navigateToLoginRequestUrl: false,
      },
      cache: {
        cacheLocation:          'localStorage',
        storeAuthStateInCookie: true,
      },
    };

    this.msal = new msal.PublicClientApplication(msalCfg);
    await this.msal.initialize();

    // Traiter le retour de redirect Microsoft
    try {
      const result = await this.msal.handleRedirectPromise();
      if (result?.account) {
        this.account = result.account;
        return;
      }
    } catch (err) {
      this.initError = err;
      console.error('[MSAL] Erreur redirect:', err);
      return;
    }

    // Session existante en cache
    const accounts = this.msal.getAllAccounts();
    if (accounts.length) this.account = accounts[0];
  }

  async login() {
    await this.msal.loginRedirect({ scopes: CONFIG.SCOPES });
  }

  async logout() {
    await this.msal.logoutRedirect({ account: this.account });
  }

  async getToken() {
    if (!this.account) throw new Error('Non authentifié');
    try {
      const r = await this.msal.acquireTokenSilent({
        scopes:  CONFIG.SCOPES,
        account: this.account,
      });
      return r.accessToken;
    } catch (err) {
      if (err instanceof msal.InteractionRequiredAuthError) {
        await this.msal.acquireTokenRedirect({ scopes: CONFIG.SCOPES, account: this.account });
      }
      throw err;
    }
  }

  isLoggedIn() { return !!this.account; }

  getUser() {
    if (!this.account) return null;
    return { name: this.account.name, email: this.account.username };
  }
}

const Auth = new BlackIPAuth();
