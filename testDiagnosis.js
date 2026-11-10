// Paste on a classic SharePoint page. Read-only JSOM/CSOM queries use POST ProcessQuery.
// Original backend CAML, including the broad project read and ZIP_Project_ID on rollouts.
// Browser identity only: this does not reproduce PHA credentials, SQL reads or Analyze().
(async () => {
  const root = location.origin; // Must match ApplicationRootUrl in the backend.
  if (window.zipPpsDiagnostics?.running) return console.error('Scan already running; wait or call zipPpsDiagnostics.stop().');
  const projectCaml = `<View><RowLimit Paged='TRUE'>5000</RowLimit></View>`;
  const rolloutCaml = `<View><ViewFields>
    <FieldRef Name='ID' /><FieldRef Name='Title' /><FieldRef Name='ZIP_HP_Code' />
    <FieldRef Name='ZIP_Project_ID' /><FieldRef Name='ZIP_ROO_ID' />
    </ViewFields><RowLimit Paged='TRUE'>5000</RowLimit></View>`;
  let stopped = false;
  const report = window.zipPpsDiagnostics = {
    started: new Date().toISOString(), root, running: true,
    mode: 'Original backend CAML through JSOM; no corrected queries',
    identity: 'Browser user; not PHA identity', queries: { projects: projectCaml, rollouts: rolloutCaml },
    projectsRead: 0, accessibleProjects: 0, requests: 0, results: [],
    stop() { stopped = true; } // In-flight read completes; no further reads start.
  };
  const read = (obj, method) => typeof obj?.[method] === 'function' ? obj[method]() : '';
  async function loadJsom() {
    if (window.SP?.ClientContext && window.SP?.CamlQuery) return;
    if (!window.SP?.SOD?.executeFunc) throw new Error('JSOM loader unavailable. Open a classic SharePoint list page and rerun.');
    await new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('Loading sp.js timed out; use a classic SharePoint list page')), 30000);
      window.SP.SOD.executeFunc('sp.js', 'SP.ClientContext', () => {
        clearTimeout(timer);
        if (window.SP?.ClientContext && window.SP?.CamlQuery) resolve();
        else reject(new Error('SharePoint JSOM did not load'));
      });
    });
  }
  function execute(context) {
    if (stopped) return Promise.reject(new Error('Scan stopped by user'));
    report.requests++;
    return new Promise((resolve, reject) => context.executeQueryAsync(resolve, (_, args) => {
      reject(Object.assign(new Error(read(args, 'get_message') || 'SharePoint query failed'), {
        errorCode: read(args, 'get_errorCode'), errorType: read(args, 'get_errorTypeName'),
        correlationId: read(args, 'get_errorTraceCorrelationId'), stackTrace: read(args, 'get_stackTrace')
      }));
    }));
  }
  async function queryPages(siteUrl, listPath, caml, visit) {
    const url = new URL(siteUrl);
    if (url.origin !== new URL(root).origin) throw new Error('Different origin discovered; browser did not query: ' + siteUrl);
    const context = new window.SP.ClientContext(siteUrl);
    let page = 0, position = null;
    try {
      const web = context.get_web();
      const list = listPath ? web.getList(url.pathname.replace(/\/$/, '') + '/' + listPath)
        : web.get_lists().getByTitle('Projekte');
      do {
        if (stopped) throw new Error('Scan stopped by user');
        if (++page > 1000) throw new Error('Diagnostic pagination limit reached; scan incomplete');
        const query = new window.SP.CamlQuery();
        query.set_viewXml(caml);
        query.set_listItemCollectionPosition(position);
        const items = list.getItems(query);
        context.load(items); // Same broad Load(items) as the backend; no Include projection.
        console.log('Query page ' + page + ': ' + siteUrl, caml);
        await execute(context);
        const enumerator = items.getEnumerator();
        while (enumerator.moveNext()) visit(enumerator.get_current());
        position = items.get_listItemCollectionPosition();
      } while (position);
    } catch (error) {
      Object.assign(error, { page, caml, requestUrl: siteUrl.replace(/\/$/, '') + '/_vti_bin/client.svc/ProcessQuery' });
      throw error;
    } finally { if (typeof context.dispose === 'function') context.dispose(); }
  }
  function strings(value) {
    if (typeof value === 'string') return value.trim() ? [value.trim()] : [];
    if (!value || typeof value !== 'object') return [];
    if (Array.isArray(value)) return value.flatMap(strings);
    // JSOM equivalents of C# Url/Description/Label/LookupValue properties.
    const values = ['get_url', 'get_description', 'get_label', 'get_lookupValue'].flatMap(method => strings(read(value, method)));
    const fallback = String(value).trim();
    if (fallback && fallback !== '[object Object]') values.push(fallback);
    return values;
  }
  function fieldText(value) {
    if (typeof value === 'string') return value.trim();
    if (value == null) return '';
    return read(value, 'get_label') || read(value, 'get_lookupValue') || read(value, 'get_url') || String(value);
  }
  function projectUrl(value) {
    const full = value.match(/https?:\/\/[^\s"'<>]+\/sites\/P-\d{4}-\d+/i);
    if (full) return full[0].replace(/\/$/, '');
    const relative = value.match(/\/sites\/P-\d{4}-\d+/i);
    if (relative) return new URL(root).origin + relative[0];
    const key = value.match(/\bP-\d{4}-\d+\b/i);
    return key ? new URL(root).origin + '/sites/' + key[0] : null;
  }
  function failure(stage, url, error, sources = []) {
    const message = error.message || String(error), lower = message.toLowerCase();
    const matchesSkipText = (lower.includes('list') && lower.includes('does not exist')) ||
      (lower.includes('liste') && lower.includes('nicht')) || lower.includes('cannot be found') || lower.includes('nicht gefunden');
    report.results.push({
      result: 'CHECK', stage, url, sources, message,
      errorCode: error.errorCode ?? '', errorType: error.errorType || '', page: error.page || '',
      caml: error.caml || '', requestUrl: error.requestUrl || '', correlationId: error.correlationId || '', stackTrace: error.stackTrace || '',
      mayMatchBackendRolloutSkip: stage === 'Original RollOutObjekte query' && matchesSkipText
    });
    console.error(stage, url, message, error);
  }
  let stage = 'Load SharePoint JSOM';
  try {
    await loadJsom();
    stage = 'Original Projekte query';
    const sites = new Map();
    await queryPages(root, null, projectCaml, item => {
      report.projectsRead++;
      const fields = item.get_fieldValues();
      if (fieldText(fields.ZIP_SC_Status).toLowerCase() !== 'accessible') return;
      report.accessibleProjects++;
      // Original discovery: scan ALL loaded fields, including descriptions and lookups.
      for (const [field, value] of Object.entries(fields)) {
        for (const text of strings(value)) {
          const url = projectUrl(text);
          if (!url) continue;
          const key = url.toLowerCase();
          if (!sites.has(key)) sites.set(key, { url, sources: [] });
          const source = 'Item ' + item.get_id() + ', field ' + field;
          if (!sites.get(key).sources.includes(source)) sites.get(key).sources.push(source);
        }
      }
    });
    report.results.push({ result: 'OK', stage, url: root, sources: [], rows: report.projectsRead });
    report.discoveredSites = sites.size;
    console.log('Accessible projects:', report.accessibleProjects, 'Discovered sites:', sites.size);
    stage = 'Original RollOutObjekte query';
    for (const site of sites.values()) {
      if (stopped) break;
      let rows = 0;
      const started = performance.now();
      try {
        await queryPages(site.url, 'Lists/RollOutObjekte', rolloutCaml, () => { rows++; });
        report.results.push({
          result: 'OK', stage, url: site.url, sources: site.sources,
          rows, seconds: Math.round((performance.now() - started) / 1000)
        });
      } catch (error) { failure(stage, site.url, error, site.sources); }
      // Diagnostic difference: continue after failed rollout sites to report all errors.
    }
  } catch (error) { failure(stage, root, error); }
  finally {
    report.running = false;
    report.stopped = stopped;
    report.finished = new Date().toISOString();
    console.table(report.results.map(({ sources, caml, stackTrace, ...row }) => ({ ...row, sources: sources.join('; ') })));
    console.log('Scan finished. A failed Projekte query prevents rollout scanning, as in the backend.');
    console.log('Original CAML, browser identity. OK does not test PHA identity, SQL, or analysis.');
    console.log('CHECK details include CAML, page, SharePoint error code and correlation ID.');
    console.log('Copy report using: copy(JSON.stringify(zipPpsDiagnostics, null, 2))');
  }
})();
