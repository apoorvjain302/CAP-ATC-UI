sap.ui.define([
  "sap/ui/core/mvc/Controller",
  "sap/ui/model/json/JSONModel",
  "sap/m/MessageToast",
  "sap/m/MessageBox",
  "sap/m/CustomListItem",
  "sap/m/StandardListItem",
  "sap/m/HBox",
  "sap/m/Button",
  "sap/m/Text",
  "sap/m/ObjectStatus",
], (Controller, JSONModel, MessageToast, MessageBox, CustomListItem, StandardListItem, HBox, Button, Text, ObjectStatus) => {
  "use strict";

  return Controller.extend("com.sap.atcanalysis.controller.Results", {

    onInit() {
      const router = this.getOwnerComponent().getRouter();
      router.getRoute("results").attachPatternMatched(this._onRouteMatched, this);
    },

    async _onRouteMatched(oEvent) {
      this._jobId = oEvent.getParameter("arguments").jobId;
      // Retrieve the CF instance affinity set during upload so all reads go to the same instance
      try { this._cfAppInstance = sessionStorage.getItem("cfAppInstance") || null; } catch (_) { this._cfAppInstance = null; }
      this._resetView();
      await this._loadResults();
    },

    _resetView() {
      ["summaryPanel","chartsPanel1","chartsPanel2","exceptionsPanel","downloadsPanel","downloadAllBtn"].forEach(id => {
        this.byId(id)?.setVisible(false);
      });
      this.byId("loadingIndicator")?.setVisible(true);
      this.byId("errorStrip")?.setVisible(false);
      this.byId("artifactsList")?.removeAllItems();
      this.byId("unmaintainedNsList")?.removeAllItems();
    },

    async _loadResults() {
      try {
        const instanceHeaders = this._cfAppInstance
          ? { "X-CF-App-Instance": this._cfAppInstance }
          : {};

        // Fetch job status — retry up to 5 times (2s apart) in case of brief DB lag after navigation
        let job;
        for (let attempt = 0; attempt < 5; attempt++) {
          const resp = await fetch(
            `/api/atc/Jobs(id='${this._jobId}')?$select=status,statusMsg,totalCount,hcaCount,s4Count,spddCount,spauCount,fitGapDeltaCount,nsFindings`,
            { headers: instanceHeaders }
          );
          const data = await resp.json();
          job = data.value?.[0] || data;
          if (job.status === "done" || job.status === "failed") break;
          // Still running — wait 2s and retry
          await new Promise(r => setTimeout(r, 2000));
        }

        if (job.status === "failed") {
          this._showError(`Analysis failed: ${job.statusMsg}`);
          return;
        }
        if (job.status !== "done") {
          this._showError(`Job is not complete (status: ${job.status}). Please wait and refresh.`);
          return;
        }

        // Fetch chart data via function
        let atcData = null;
        try {
          const cdResp = await fetch(
            `/api/atc/Jobs(id='${this._jobId}')/ATCService.getChartData()`,
            { headers: instanceHeaders }
          );
          const cdJson = await cdResp.json();
          const raw = cdJson.value;
          if (raw && typeof raw === "string" && raw !== "null") {
            atcData = JSON.parse(raw);
          } else if (raw && typeof raw === "object") {
            atcData = raw;
          }
        } catch (chartErr) {
          console.warn("Chart data unavailable:", chartErr.message);
        }

        // Fetch artifacts
        const artResp = await fetch(
          `/api/atc/Artifacts?$filter=jobId eq '${this._jobId}'&$select=id,role,filename,mimeType,size`,
          { headers: instanceHeaders }
        );
        const artData = await artResp.json();
        const artifacts = artData.value || [];

        if (!artifacts.length) {
          this._showError("No output files found for this job. The server may have restarted and lost the session data. Please run the analysis again.");
          return;
        }

        this.byId("loadingIndicator").setVisible(false);
        this._renderSummary(job);
        this._renderCharts(atcData);
        this._renderArtifacts(artifacts);
        this._renderExceptions(job);

        this.byId("summaryPanel").setVisible(true);
        this.byId("chartsPanel1").setVisible(!!atcData);
        this.byId("chartsPanel2").setVisible(!!atcData);
        this.byId("downloadsPanel").setVisible(true);
        this.byId("downloadAllBtn").setVisible(true);

      } catch (err) {
        this._showError(`Failed to load results: ${err.message}`);
      }
    },

    _renderSummary(job) {
      this.byId("totalCountTile").setValue(String(job.totalCount || 0));
      this.byId("hcaCountTile").setValue(String(job.hcaCount   || 0));
      this.byId("s4CountTile").setValue(String(job.s4Count     || 0));
      this.byId("spddTile").setValue(String(job.spddCount || 0));
      this.byId("spauTile").setValue(String(job.spauCount || 0));
    },

    _renderCharts(atcData) {
      if (!atcData) return;

      // Slide 5 — remediation categories
      const slide5Model = new JSONModel({
        items: (atcData.slide5Categories || []).map((cat, i) => ({
          category: cat,
          count:    (atcData.slide5Values || [])[i] || 0,
        })),
      });
      this.byId("slide5Dataset").setModel(slide5Model);
      this.byId("slide5Dataset").bindAggregation("data", { path: "/items" });
      _applyVizProps(this.byId("slide5Chart"), "Remediation Categories");

      // Slide 4 — object types
      const slide4Model = new JSONModel({
        items: (atcData.uniqueObjTypes || []).map((type, i) => ({
          type,
          count: (atcData.uniqueObjCounts || [])[i] || 0,
        })),
      });
      this.byId("slide4Dataset").setModel(slide4Model);
      this.byId("slide4Dataset").bindAggregation("data", { path: "/items" });
      _applyVizProps(this.byId("slide4Chart"), "Object Types");

      // HCA waterfall
      const hcaValues = atcData.chart7Values || [];
      const hcaTotal  = hcaValues.reduce((s, v) => s + (v || 0), 0);
      const hcaModel = new JSONModel({
        items: (atcData.chart7Categories || []).map((cat, i) => ({
          category: cat,
          count:    hcaValues[i] || 0,
        })),
      });
      this.byId("hcaDataset").setModel(hcaModel);
      this.byId("hcaDataset").bindAggregation("data", { path: "/items" });
      this.byId("hcaChartLabel").setText(`HCA Impact Breakdown — ${hcaTotal} findings across all categories`);
      _applyVizProps(this.byId("hcaChart"), "HCA Breakdown");

      // S/4H waterfall
      const s4hValues = atcData.chart8Values || [];
      const s4hTotal  = s4hValues.reduce((s, v) => s + (v || 0), 0);
      const s4hModel = new JSONModel({
        items: (atcData.chart8Categories || []).map((cat, i) => ({
          category: cat,
          count:    s4hValues[i] || 0,
        })),
      });
      this.byId("s4hDataset").setModel(s4hModel);
      this.byId("s4hDataset").bindAggregation("data", { path: "/items" });
      this.byId("s4hChartLabel").setText(`S/4HANA Impact Breakdown — ${s4hTotal} findings across all categories`);
      _applyVizProps(this.byId("s4hChart"), "S/4HANA Breakdown");
    },

    _renderArtifacts(artifacts) {
      const list = this.byId("artifactsList");
      list.removeAllItems();

      const filtered = artifacts;

      if (!filtered.length) {
        list.addItem(new CustomListItem({
          content: [new Text({ text: "No output files generated yet." })]
        }));
        return;
      }

      const roleLabels = {
        atc_result: "ATC Classified Results (XLSX)",
        tua:        "TUA Analysis (XLSX)",
        pptx:       "Presentation (PPTX)",
        estimation: "Effort Estimation (XLSX)",
      };

      filtered.forEach(art => {
        // Per-chunk API artifacts have roles like atc_result_chunk_1, atc_result_chunk_2, ...
        let label;
        const chunkMatch = art.role && art.role.match(/^atc_result_chunk_(\d+)$/);
        if (chunkMatch) {
          label = `API Result — Chunk ${chunkMatch[1]} (XLS)`;
        } else {
          label = roleLabels[art.role] || art.filename;
        }
        const sizeText = _formatSize(art.size);
        const item = new CustomListItem({
          content: [
            new HBox({
              alignItems: "Center",
              justifyContent: "SpaceBetween",
              items: [
                new HBox({
                  alignItems: "Center",
                  items: [
                    new ObjectStatus({
                      icon:  _roleIcon(art.role),
                      state: "Success",
                      class: "sapUiSmallMarginEnd",
                    }),
                    new Text({ text: `${label} — ${art.filename} (${sizeText})` }),
                  ],
                }),
                new Button({
                  text:  "Download",
                  icon:  "sap-icon://download",
                  type:  "Transparent",
                  press: () => { _downloadArtifact(art.id, art.filename, this._cfAppInstance); },
                }),
              ],
            }),
          ],
        });
        list.addItem(item);
      });
    },

    _renderExceptions(job) {
      const fitGapDelta = job.fitGapDeltaCount || 0;
      let unmaintained = [];
      if (job.nsFindings) {
        try {
          const f = JSON.parse(job.nsFindings);
          unmaintained = f.unmaintained || [];
        } catch (_) {}
      }

      // Always show exceptions panel — it has Fit Gap Delta even for ATC-only jobs
      this.byId("exceptionFitGapDelta").setText(String(fitGapDelta));

      const unmaintainedList = this.byId("unmaintainedNsList");
      unmaintained.forEach(ns => {
        unmaintainedList.addItem(new StandardListItem({ title: ns, icon: "sap-icon://warning" }));
      });

      this.byId("exceptionsPanel").setVisible(true);
    },

    async onDownloadAll() {
      if (!this._jobId) return;
      try {
        const headers = this._cfAppInstance ? { "X-CF-App-Instance": this._cfAppInstance } : {};
        const resp = await fetch(`/atc/downloadAll/${this._jobId}`, { headers });
        if (!resp.ok) {
          const errText = await resp.text();
          this._showError(`Download failed: ${errText.slice(0, 200)}`);
          return;
        }
        const blob = await resp.blob();
        const url  = URL.createObjectURL(blob);
        const disp = resp.headers.get("Content-Disposition") || "";
        const nameMatch = disp.match(/filename="?([^"]+)"?/);
        const filename  = nameMatch ? nameMatch[1] : "ATC_Results.zip";
        const link = document.createElement("a");
        link.href     = url;
        link.download = filename;
        link.click();
        URL.revokeObjectURL(url);
      } catch (err) {
        this._showError(`Download error: ${err.message}`);
      }
    },

    onNavBack() {
      MessageBox.confirm(
        "Going back will clear all uploaded files and start a new analysis. Continue?",
        {
          title: "Leave Results Page",
          actions: [MessageBox.Action.YES, MessageBox.Action.NO],
          emphasizedAction: MessageBox.Action.NO,
          onClose: (action) => {
            if (action !== MessageBox.Action.YES) return;
            this.getOwnerComponent().getRouter().navTo("home");
          },
        }
      );
    },

    _showError(msg) {
      this.byId("loadingIndicator").setVisible(false);
      const strip = this.byId("errorStrip");
      strip.setText(msg).setVisible(true);
    },
  });
});

// ── Helpers ──────────────────────────────────────────────────────────────────
function _applyVizProps(vizFrame, title) {
  if (!vizFrame) return;
  vizFrame.setVizProperties({
    title:  { visible: false },
    legend: { visible: false },
    categoryAxis: { title: { visible: false } },
    valueAxis:    { title: { visible: false } },
    plotArea:     { dataLabel: { visible: true } },
  });
}

async function _downloadArtifact(artifactId, filename, cfAppInstance) {
  try {
    const headers = cfAppInstance ? { "X-CF-App-Instance": cfAppInstance } : {};
    const resp = await fetch(`/atc/download/${artifactId}`, { headers });
    if (!resp.ok) { console.error("Download failed", resp.status); return; }
    const blob = await resp.blob();
    const url  = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href     = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
  } catch (err) {
    console.error("Download error:", err.message);
  }
}

function _roleIcon(role) {
  if (role && role.startsWith("atc_result_chunk_")) return "sap-icon://document";
  const icons = {
    atc_result: "sap-icon://document",
    tua:        "sap-icon://table-chart",
    pptx:       "sap-icon://present",
    estimation: "sap-icon://budget",
  };
  return icons[role] || "sap-icon://document";
}

function _formatSize(bytes) {
  if (!bytes) return "0 B";
  const kb = bytes / 1024;
  return kb < 1024 ? `${kb.toFixed(1)} KB` : `${(kb/1024).toFixed(1)} MB`;
}
