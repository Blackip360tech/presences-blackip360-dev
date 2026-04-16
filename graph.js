// BlackIP360 Présences — Microsoft Graph API
class GraphAPI {
  constructor() {
    this._siteId = null;
    this._listId = null;
  }

  // ── Requête générique authentifiée ────────────────────────────────────────
  async _call(path, options = {}) {
    const token = await Auth.getToken();
    const res   = await fetch(CONFIG.GRAPH_BASE + path, {
      ...options,
      headers: {
        Authorization:  `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...options.headers,
      },
    });

    if (res.status === 204) return null;

    const data = await res.json();
    if (!res.ok) throw new Error(data.error?.message || `HTTP ${res.status}`);
    return data;
  }

  // ── Identifiants SharePoint (mis en cache) ────────────────────────────────
  async _siteIdCached() {
    if (this._siteId) return this._siteId;
    const d = await this._call(
      `/sites/${CONFIG.SHAREPOINT_HOST}:${CONFIG.SHAREPOINT_SITE_PATH}`
    );
    this._siteId = d.id;
    return d.id;
  }

  async _listIdCached() {
    if (this._listId) return this._listId;
    const sid = await this._siteIdCached();
    const d   = await this._call(`/sites/${sid}/lists/${CONFIG.SHAREPOINT_LIST}`);
    this._listId = d.id;
    return d.id;
  }

  // ── Profil de l'utilisateur connecté ─────────────────────────────────────
  async getProfile() {
    return this._call('/me?$select=displayName,mail,jobTitle,department');
  }

  // ── Toutes les présences (500 max, triées par heure desc) ─────────────────
  async getAllPresences() {
    const sid = await this._siteIdCached();
    const lid = await this._listIdCached();
    const url = `/sites/${sid}/lists/${lid}/items`
      + `?$expand=fields`
      + `&$orderby=fields/HeurePointage desc`
      + `&$top=500`;
    const d = await this._call(url);
    return (d.value || []).map(i => ({ id: i.id, ...i.fields }));
  }

  // ── Présences d'un employé spécifique ─────────────────────────────────────
  async getMyPresences(email) {
    const sid    = await this._siteIdCached();
    const lid    = await this._listIdCached();
    const filter = encodeURIComponent(`fields/EmployeEmail eq '${email}'`);
    const url    = `/sites/${sid}/lists/${lid}/items`
      + `?$filter=${filter}`
      + `&$expand=fields`
      + `&$orderby=fields/HeurePointage desc`
      + `&$top=100`;
    const d = await this._call(url);
    return (d.value || []).map(i => ({ id: i.id, ...i.fields }));
  }

  // ── Statut actuel de chaque employé (entrée la plus récente par email) ────
  async getCurrentStatuses() {
    const all = await this.getAllPresences();
    const map = {};
    for (const p of all) {
      const key = p.EmployeEmail;
      if (!map[key] || new Date(p.HeurePointage) > new Date(map[key].HeurePointage)) {
        map[key] = p;
      }
    }
    return Object.values(map);
  }

  // ── Créer un pointage ─────────────────────────────────────────────────────
  async pointage({ nom, email, departement, statut, notes }) {
    const sid  = await this._siteIdCached();
    const lid  = await this._listIdCached();
    return this._call(`/sites/${sid}/lists/${lid}/items`, {
      method: 'POST',
      body: JSON.stringify({
        fields: {
          EmployeNom:    nom,
          EmployeEmail:  email,
          Departement:   departement,
          StatutActuel:  statut,
          HeurePointage: new Date().toISOString(),
          Notes:         notes || '',
        },
      }),
    });
  }
}

const Graph = new GraphAPI();
