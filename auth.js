// BlackIP360 Présences — Authentification MSAL.js (redirect-only)
// Redirect uniquement : compatible Teams, AuthPoint, navigateurs stricts
class BlackIPAuth {
  constructor() {
    this.msal    = null;
    this.account = null;
  }

  async init() {
    const msalCfg = {
      auth: {
        clientId:              CONFIG.CLIENT_ID,
        authority:             `https://login.microsoftonline.com/${CONFIG.TENANT_ID}`,
        redirectUri:           CONFIG.APP_URL + '/',
        postLogoutRedirectUri: CONFIG.APP_URL + '/',
        navigateToLoginRequestUrl: true,
      },
      cache: {
        cacheLocation:          'localStorage',  // persist entre les redirects
        storeAuthStateInCookie: true,            // requis pour certains navigateurs strict
      },
    };

    this.msal = new msal.PublicClientApplication(msalCfg);
    await this.msal.initialize();

    // Intercepter le retour de redirect Microsoft
    try {
      const result = await this.msal.handleRedirectPromise();
      if (result?.account) {
        this.account = result.account;
        return;
      }
    } catch (err) {
      console.error('[MSAL] handleRedirectPromise error:', err);
    }

    // Reprendre la session existante si présente
    const accounts = this.msal.getAllAccounts();
    if (accounts.length) this.account = accounts[0];
  }

  async login() {
    await this.msal.loginRedirect({
      scopes: CONFIG.SCOPES,
      prompt: 'select_account',
    });
    // La page se redirige vers Microsoft — le retour est géré dans init()
  }

  async logout() {
    await this.msal.logoutRedirect({
      account:               this.account,
      postLogoutRedirectUri: CONFIG.APP_URL + '/',
    });
  }

  async getToken() {
    if (!this.account) throw new Error('Non authentifié — veuillez vous connecter.');
    try {
      const r = await this.msal.acquireTokenSilent({
        scopes:  CONFIG.SCOPES,
        account: this.account,
      });
      return r.accessToken;
    } catch (err) {
      if (err instanceof msal.InteractionRequiredAuthError) {
        // Token silencieux impossible → redirect
        await this.msal.acquireTokenRedirect({
          scopes:  CONFIG.SCOPES,
          account: this.account,
        });
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
