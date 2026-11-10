// Paste this entire file into DevTools Console on the ABN ZIP administration page.
// Runs the deployed Preview(15), including SQL, SharePoint reads and C# Analyze().
// GET requests only; no upload/apply. Backend execution stops at its first unhandled error.
// Full exceptions require access to the existing PHA Log Viewer API.
(async () => {
    const baseUrl = 'https://zip-pha.cloud.abn-on.f-i.de';
    const batchId = 15;
    if (![baseUrl, 'https://zip.cloud.abn-on.f-i.de'].includes(location.origin)) {
        console.error('Open the ABN ZIP administration page before running this script.');
        return;
    }
    if (window.zipPpsPreview15?.running) {
        console.warn('Preview is already running.');
        return;
    }

    const started = performance.now();
    const report = window.zipPpsPreview15 = {
        batchId, baseUrl, startedAt: new Date().toISOString(), running: true,
        phase: 'Preview', requests: [], logDetails: [], correlationId: null
    };
    const errorInfo = error => ({ name: error.name, message: error.message, stack: error.stack });

    async function get(path) {
        const request = { url: baseUrl + path, startedAt: new Date().toISOString() };
        report.requests.push(request);
        const start = performance.now();
        try {
            // Same authentication as PpsMasterDataAdminService.preview(). No custom headers/preflight.
            const response = await fetch(request.url, { method: 'GET', credentials: 'include', cache: 'no-store' });
            Object.assign(request, {
                status: response.status, statusText: response.statusText, ok: response.ok,
                responseUrl: response.url, redirected: response.redirected,
                headers: Object.fromEntries(response.headers.entries())
            });
            request.rawBody = await response.text();
            try { request.body = JSON.parse(request.rawBody); }
            catch { request.body = request.rawBody; }
            console.log('GET', request.url, 'HTTP', request.status, request.body);
            return request;
        } catch (error) {
            request.error = errorInfo(error);
            throw error;
        } finally {
            request.elapsedMs = Math.round(performance.now() - start);
        }
    }

    // Can be rerun separately if the log entry has not yet been flushed to SQL.
    report.loadLogs = async () => {
        if (!report.correlationId) {
            console.warn('No correlation ID was returned. Inspect the saved HTTP response and Network tab.');
            return;
        }
        try {
            const result = await get('/api/logs/pha?correlationId=' + encodeURIComponent(report.correlationId) + '&pageSize=20');
            report.logQuery = result;
            if (!result.ok || !Array.isArray(result.body?.Items)) {
                console.warn('Log lookup unavailable. HTTP 401/403 requires Log Viewer access; 404 may mean this API is not deployed. The original preview response is saved.');
                return;
            }
            if (!result.body.Items.length) {
                console.warn('No matching logs yet. Retry later: await zipPpsPreview15.loadLogs()');
                return;
            }
            console.table(result.body.Items);
            if (result.body.HasMore) console.warn('More than 20 matching logs exist; use the PHA Log Viewer for the remaining entries.');
            report.logDetails = [];
            for (const entry of result.body.Items) {
                if (!Number.isInteger(entry.Id) || entry.Id < 1) continue;
                // List responses omit Exception and truncate Message; fetch the detail endpoint.
                const detail = await get('/api/logs/pha/' + entry.Id);
                if (!detail.ok) continue;
                report.logDetails.push(detail.body);
                console.log('PHA log ' + entry.Id + ':', detail.body.Message);
                console.log('Full exception / inner exceptions / stack trace:', detail.body.Exception || '(No exception recorded)');
            }
        } catch (error) {
            report.logLookupError = errorInfo(error);
            console.error('Log lookup failed; the original preview response is preserved.', error);
        }
    };

    console.log('Running complete ABN preview for batch 15:', baseUrl + '/api/pps-masterdata-import/preview/15');
    console.log('Server sequence: ReadConfiguration -> ReadStagedRows -> CreateSharePointContext -> ReadMasterData -> ReadProjectAndRolloutUsages -> Analyze -> DisposeSharePointContext');
    const heartbeat = setInterval(() => {
        console.log(report.phase + ' pending after ' + Math.round((performance.now() - started) / 1000) + ' seconds. The API does not stream internal step progress.');
    }, 15000);
    try {
        const response = await get('/api/pps-masterdata-import/preview/' + batchId);
        report.previewResponse = response;
        if (!response.ok) {
            report.outcome = 'HTTP error';
            const match = response.rawBody.match(/(?:Korrelations-ID|Correlation\s*ID)\s*:\s*([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i);
            report.correlationId = match ? match[1] : null;
            console.error('Preview failed. HTTP', response.status, 'Correlation ID:', report.correlationId, response.body);
            report.phase = 'Log lookup';
            await report.loadLogs();
        } else if (response.body?.BatchId !== batchId || !Array.isArray(response.body?.Rows)) {
            report.outcome = 'Unexpected response';
            console.error('HTTP succeeded but did not return a batch 15 preview. Check for a login page or redirect.', response);
        } else {
            report.outcome = 'Success';
            report.preview = response.body;
            const { Rows, Headers, ...summary } = response.body;
            console.table([summary]);
            console.table(Rows.map(row => ({
                RowNumber: row.RowNumber, HpCode: row.HpCode, Action: row.Action,
                IsBlocked: row.IsBlocked, BlockReason: row.BlockReason,
                ChangedFields: (row.ChangedFields || []).join(', '),
                Warnings: (row.Warnings || []).join(' | ')
            })));
            console.log('Complete preview, including cells, usages, hints and propagation:', report.preview);
        }
    } catch (error) {
        report.outcome = 'Browser/request error';
        report.error = errorInfo(error);
        console.error('Browser could not complete/read the request.', error);
        console.warn('Check Network for authentication, CORS, TLS or proxy errors. If CORS hides the response, open the ABN PHA origin in another tab and run there. A browser error does not prove the server stopped executing.');
    } finally {
        clearInterval(heartbeat);
        report.running = false;
        report.finishedAt = new Date().toISOString();
        report.elapsedMs = Math.round(performance.now() - started);
        console.log('Finished:', report.outcome, report);
        console.log('Copy the report: copy(JSON.stringify(zipPpsPreview15, null, 2))');
    }
})();
