fetch("https://zip-pha.cloud.abn-on.f-i.de/api/pps-masterdata-import/current-version", {
    credentials: "include"
}).then(async r => console.log(r.status, await r.text()));