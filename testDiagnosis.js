// Paste this entire file into DevTools Console on the target SharePoint page.

// GET requests only. Uses your browser identity, NOT the PHA backend identity.

// REST field representations/queries differ from CSOM: this is a diagnostic,

// not an exact replay of Preview(). No upload, apply, or list updates are made.

(async () => {

  // The Projekte list is on the root web, even when this runs on an admin page.

  const root = location.origin;

  if (window.zipPpsDiagnostics?.running) {

    console.error("A scan is already running. Wait or call zipPpsDiagnostics.stop().");

    return;

  }

  let stopped = false;

  let activeController;

  const report = window.zipPpsDiagnostics = {

    started: new Date().toISOString(), root, running: true,

    identity: "Current browser user; not PHA service identity",

    projectsRead: 0, accessibleProjects: 0, requests: 0,

    results: [],

    stop() { stopped = true; activeController?.abort(); }

  };

  const literal = value => "'" + encodeURIComponent(value.replace(/'/g, "''")).replace(/'/g, "%27") + "'";

  async function get(url) {

    if (stopped) throw new Error("Scan stopped by user");

    const target = new URL(url, root);

    if (target.origin !== root) throw new Error("Different host; not requested: " + target.origin);

    const controller = new AbortController();

    activeController = controller;

    const timer = setTimeout(() => controller.abort(), 60000);

    report.requests++;

    try {

      const response = await fetch(target.href, {

        method: "GET", credentials: "include", redirect: "error",

        headers: { Accept: "application/json;odata=verbose" }, signal: controller.signal

      });

      const body = await response.text();

      let data;

      try { data = JSON.parse(body); } catch (_) { /* login/error HTML */ }

      if (!response.ok || !data || data.error || data["odata.error"]) {

        const error = data?.error || data?.["odata.error"];

        const detail = error?.message?.value || error?.message ||

          (data ? "Unexpected JSON response" : "Non-JSON response (possibly a login or IIS error page)");

        throw Object.assign(new Error(String(detail).slice(0, 1500)), {

          status: response.status,

          correlationId: response.headers.get("SPRequestGuid") || "",

          requestUrl: target.href

        });

      }

      return data;

    } catch (error) {

      error.requestUrl = error.requestUrl || target.href;

      if (error.name === "AbortError") {

        throw Object.assign(new Error(stopped ? "Scan stopped by user" : "Diagnostic request exceeded 60 seconds"), {

          requestUrl: target.href

        });

      }

      throw error;

    } finally { clearTimeout(timer); activeController = null; }

  }

  async function eachPage(url, visit) {

    const visited = new Set();

    while (url) {

      if (visited.has(url) || visited.size >= 1000) throw new Error("Pagination safety limit reached; scan incomplete");

      visited.add(url);

      const data = await get(url);

      const rows = data.d?.results || data.value;

      if (!Array.isArray(rows)) throw new Error("Expected a SharePoint collection response");

      visit(rows);

      url = data.d?.__next || data["@odata.nextLink"] || data["odata.nextLink"];

    }

  }

  function strings(value) {

    if (typeof value === "string") return [value];

    if (!value || typeof value !== "object") return [];

    return Object.entries(value).filter(([key]) => !key.startsWith("__"))

      .flatMap(([, child]) => strings(child));

  }

  function projectUrl(value) {

    const full = value.match(/https?:\/\/[^\s"'<>]+\/sites\/P-\d{4}-\d+/i);

    if (full) return full[0].replace(/\/$/, "");

    const relative = value.match(/\/sites\/P-\d{4}-\d+/i);

    if (relative) return root + relative[0];

    const key = value.match(/\bP-\d{4}-\d+\b/i);

    return key ? root + "/sites/" + key[0] : null;

  }

  function failure(stage, url, error, sources = []) {

    report.results.push({

      result: "CHECK", stage, url, sources,

      status: error.status || "", message: error.message,

      requestUrl: error.requestUrl || "", correlationId: error.correlationId || ""

    });

    console.warn(stage, url, error.message);

  }

  let stage = "Read Projekte columns";

  try {

    const projectsApi = root + "/_api/web/lists/getbytitle('Projekte')";

    const columns = [];

    await eachPage(projectsApi + "/fields?$select=InternalName,EntityPropertyName&$top=5000", rows => columns.push(...rows));

    const required = ["ID", "ZIP_SC_Status"];

    const missing = required.filter(name => !columns.some(field => field.InternalName === name));

    if (missing.length) throw new Error("Missing project fields: " + missing.join(", ") +

      ". Cannot identify Accessible projects; scan incomplete.");



    // Do not request all fields: Projekte has enough lookup/person fields to

    // exceed SharePoint's lookup-column threshold. A smaller $top alone does

    // not fix that. These provisioned fields are scalar text/choice fields.

    const projectNames = ["ID", "Title", "ZIP_SC_Status", "ZIP_Project_ID", "ZIP_ProjectSiteUrl"];

    const projectFields = columns.filter(field => projectNames.includes(field.InternalName));

    const propertyName = field => field.EntityPropertyName || (field.InternalName === "ID" ? "Id" : field.InternalName);

    const statusProperty = propertyName(projectFields.find(field => field.InternalName === "ZIP_SC_Status"));

    const idProperty = propertyName(projectFields.find(field => field.InternalName === "ID"));

    report.projectFields = projectFields.map(field => field.InternalName);

    report.discoveryScope = "Project site URL, project ID and title only; links in other fields are not scanned";

    console.log(report.discoveryScope);



    stage = "Read Projekte rows";

    const sites = new Map();

    await eachPage(projectsApi + "/items?$select=" + projectFields.map(propertyName).join(",") + "&$top=500", rows => {

      report.projectsRead += rows.length;

      for (const row of rows) {

        if (String(row[statusProperty] || "").trim().toLowerCase() !== "accessible") continue;

        report.accessibleProjects++;

        for (const [field, value] of Object.entries(row)) {

          if (field.startsWith("__")) continue;

          for (const text of strings(value)) {

            const url = projectUrl(text.trim());

            if (!url) continue;

            const key = url.toLowerCase();

            if (!sites.has(key)) sites.set(key, { url, sources: [] });

            const source = "Item " + row[idProperty] + ", field " + field;

            if (!sites.get(key).sources.includes(source)) sites.get(key).sources.push(source);

          }

        }

      }

      console.log("Projects read:", report.projectsRead, "Discovered sites:", sites.size);

    });

    report.discoveredSites = sites.size;

    let index = 0;

    for (const site of sites.values()) {

      if (stopped) break;

      console.log("Checking site " + (++index) + "/" + sites.size + ": " + site.url);

      let siteStage = "Check discovered URL";

      const started = performance.now();

      try {

        const url = new URL(site.url);

        if (url.origin !== root) throw new Error("Different host found in project data; not requested by this browser scan");

        const listPath = url.pathname.replace(/\/$/, "") + "/Lists/RollOutObjekte";

        const listApi = site.url + "/_api/web/GetList(" + literal(listPath) + ")";

        siteStage = "Read RollOutObjekte columns";

        const fields = [];

        await eachPage(listApi + "/fields?$select=InternalName,EntityPropertyName&$top=5000", rows => fields.push(...rows));

        const names = ["ID", "Title", "ZIP_HP_Code", "ZIP_ROO_ZugeordnetesProjektID", "ZIP_ROO_ID"];

        const absent = names.filter(name => !fields.some(field => field.InternalName === name));

        if (absent.length) throw new Error("Missing rollout fields: " + absent.join(", "));

        const select = names.map(name => propertyName(fields.find(field => field.InternalName === name)));

        siteStage = "Read RollOutObjekte rows";

        let count = 0;

        await eachPage(listApi + "/items?$select=" + select.join(",") + "&$top=500", rows => { count += rows.length; });

        report.results.push({

          result: "OK", stage: siteStage, url: site.url, sources: site.sources,

          rows: count, seconds: Math.round((performance.now() - started) / 1000)

        });

      } catch (error) { failure(siteStage, site.url, error, site.sources); }

    }

  } catch (error) { failure(stage, root, error); }

  finally {

    report.running = false;

    report.stopped = stopped;

    report.finished = new Date().toISOString();

    console.table(report.results.map(({ sources, ...row }) => ({ ...row, sources: sources.join("; ") })));

    console.log("Scan finished. CHECK entries are suspects, not proof of the PHA failure.");

    console.log("REST differs from CSOM; OK does not test PHA identity, SQL staging reads, or analysis.");

    console.log("A missing list may be skipped by the backend. Network failures may be browser-only restrictions.");

    console.log("Copy report using: copy(JSON.stringify(zipPpsDiagnostics, null, 2))");

  }

})();

