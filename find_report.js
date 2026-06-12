const { PowerBiClient, getAccessToken } = require("@pbi-lens/core");
const TARGET_REPORT = "20f08860-86b8-402b-889f-c325399de2fa";
const TARGET_DATASET = "3ef0956e-6782-4c5f-967a-ebc2d4a7aabb";
(async () => {
  const token = await getAccessToken();
  const client = new PowerBiClient(token);
  const wss = await client.listWorkspaces();
  for (const ws of wss) {
    let reports;
    try { reports = await client.listReports(ws.id); }
    catch (e) { continue; }
    for (const r of reports) {
      if (r.id === TARGET_REPORT || r.datasetId === TARGET_DATASET || /M(ä|ae)ngel.?Statistik|MPS2/i.test(r.name || "")) {
        console.log(JSON.stringify({ workspace: ws.name, workspaceId: ws.id, report: r.name, reportId: r.id, datasetId: r.datasetId }));
      }
    }
  }
  console.log("DONE");
})().catch(e => { console.error("ERR", e.message); process.exit(1); });
