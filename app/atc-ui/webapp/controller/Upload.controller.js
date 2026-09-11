sap.ui.define([
  "sap/ui/core/mvc/Controller",
  "sap/m/MessageToast",
  "sap/m/MessageBox",
], (Controller, MessageToast, MessageBox) => {
  "use strict";

  return Controller.extend("com.sap.atcanalysis.controller.Upload", {

    onInit() {
      this._fileIds = {
        atc:       null,
        clone:     null,
        smodilog:  null,
        trnspacet: null,
      };
      // Store actual File objects from change events
      this._files = {
        atc:       null,
        clone:     null,
        smodilog:  null,
        trnspacet: null,
      };
      // CF instance affinity: captured from first upload response.
      // Sent as X-CF-App-Instance on all subsequent requests to guarantee
      // all data (files, job, runAnalysis) lands on the same CF instance/SQLite DB.
      this._cfAppInstance = null;
      // Default analysis mode
      this._analysisMode = "atc";
      // Default: use App Logic (lower memory, no Cloud Connector needed)
      this._useAppLogic = true;
      // Default migration type: Conversion (ECC → S/4HANA)
      this._migType = "conversion";
      this.getOwnerComponent().getRouter()
        .getRoute("home")
        .attachPatternMatched(this._onRouteMatched, this);
    },

    _onRouteMatched() {
      this.resetView();
    },

    resetView() {
      // Clear file references
      this._fileIds = { atc: null, clone: null, smodilog: null, trnspacet: null };
      this._files   = { atc: null, clone: null, smodilog: null, trnspacet: null };
      this._cfAppInstance = null;
      this._analysisMode = "atc";
      this._useAppLogic = true;
      this._migType = "conversion";
      this.byId("customerInput")?.setValue("");
      this.byId("atcFileUploader")?.clear?.();
      this.byId("cloneFileUploader")?.clear?.();
      this.byId("smodilogUploader")?.clear?.();
      this.byId("trnspacetUploader")?.clear?.();
      this.byId("atcFileStatus")?.setText("No file selected").setState("None");
      this.byId("cloneFileStatus")?.setText("No file selected").setState("None");
      this.byId("smodilogStatus")?.setText("No file selected").setState("None");
      this.byId("trnspacetStatus")?.setText("No file selected").setState("None");
      this.byId("analysisModeBtn")?.setSelectedKey("atc");

      this.byId("migrationTypeBtn")?.setSelectedItem(this.byId("migConversionItem"));
      this.byId("errorStrip")?.setVisible(false);
      this._applyModeVisibility();
      this._setRunning(false);
      this._updateRunButton();
    },

    // ── Analysis mode change ──────────────────────────────────────────────────
    onModeChange(oEvent) {
      this._analysisMode = oEvent.getParameter("item").getKey();
      this._applyModeVisibility();
      this._updateRunButton();
    },

    _applyModeVisibility() {
      const m = this._analysisMode;
      const showAtc   = m !== "tua_only";
      const showClone = m === "atc_clone" || m === "atc_tua_clone";
      const showTua   = m === "atc_tua" || m === "atc_tua_clone" || m === "tua_only";
      this.byId("atcUploaderPanel")?.setVisible(showAtc);
      this.byId("cloneUploaderPanel")?.setVisible(showClone);
      this.byId("tuaUploaderPanel")?.setVisible(showTua);
    },

    // ── Customer name change ──────────────────────────────────────────────────
    onCustomerChange() {
      this._updateRunButton();
    },

    // ── Migration type toggle ──────────────────────────────────────────────────
    onMigrationTypeChange(oEvent) {
      this._migType = oEvent.getParameter("item").getKey(); // "conversion" | "upgrade"
    },

    // ── File selection callbacks ───────────────────────────────────────────────
    onAtcFileChange(oEvent) {
      const file = oEvent.getParameter("files")[0];
      const status = this.byId("atcFileStatus");
      if (file) {
        this._files.atc = file;
        status.setText(file.name).setState("Success");
      } else {
        this._files.atc = null;
        this._fileIds.atc = null;
        status.setText("No file selected").setState("None");
      }
      this._updateRunButton();
    },

    onOptionalFileChange(oEvent) {
      const file = oEvent.getParameter("files")[0];
      const uid  = oEvent.getSource().getId();
      const role = uid.includes("clone")     ? "clone"
                 : uid.includes("smodilog")  ? "smodilog"
                 : uid.includes("trnspacet") ? "trnspacet"
                 : null;
      if (role) {
        this._files[role] = file || null;
        const statusMap = {
          clone:     "cloneFileStatus",
          smodilog:  "smodilogStatus",
          trnspacet: "trnspacetStatus",
        };
        const statusId = statusMap[role];
        if (statusId) {
          this.byId(statusId)
            ?.setText(file ? file.name : "No file selected")
            .setState(file ? "Success" : "None");
        }
        this._updateRunButton();
      }
    },

    // ── Run Analysis ──────────────────────────────────────────────────────────
    async onRunAnalysis() {
      const customer = this.byId("customerInput").getValue().trim();
      const m = this._analysisMode;

      if (!customer) {
        MessageBox.error("Please enter a customer name.");
        return;
      }
      if (m !== "tua_only" && !this._files.atc) {
        MessageBox.error("Please select the ATC file before running analysis.");
        return;
      }

      this._setRunning(true, "Uploading files...");
      this.byId("errorStrip").setVisible(false);

      try {
        // Upload all selected files
        await this._uploadAllFiles();

        if (m !== "tua_only" && !this._fileIds.atc) {
          throw new Error("ATC file upload did not return a file ID.");
        }

        // Create job — must go to the same instance that received the file uploads
        this._setRunning(true, "Creating analysis job...");
        const instanceHeaders = this._cfAppInstance
          ? { "Content-Type": "application/json", "X-CF-App-Instance": this._cfAppInstance }
          : { "Content-Type": "application/json" };
        const createResp = await fetch("/atc/jobs/create", {
          method:  "POST",
          headers: instanceHeaders,
          body:    JSON.stringify({
            customer,
            analysisMode: this._analysisMode,
            atcFileId:    this._fileIds.atc      || null,
            cloneFileId:  this._fileIds.clone    || null,
            smodilogId:   this._fileIds.smodilog || null,
            trnspacetId:  this._fileIds.trnspacet|| null,
            nsOwnerId:    null,
            useAppLogic:  this._useAppLogic,
            migType:      this._migType,
          }),
        });
        const createResult = await createResp.json();
        if (createResult.error) throw new Error(createResult.error);

        const jobId = createResult.jobId;
        this._setRunning(true, "Running analysis...");

        // Trigger analysis — must reach the same instance where job and files live
        const runHeaders = { "Content-Type": "application/json" };
        if (this._cfAppInstance) runHeaders["X-CF-App-Instance"] = this._cfAppInstance;
        const runResp = await fetch(
          `/api/atc/Jobs(id='${jobId}')/ATCService.runAnalysis`,
          {
            method:  "POST",
            headers: runHeaders,
            body:    JSON.stringify({}),
          },
        );
        if (!runResp.ok) {
          const errText = await runResp.text();
          throw new Error(`Run action failed: ${errText.slice(0, 300)}`);
        }

        // Poll for completion
        this._pollJob(jobId);

      } catch (err) {
        this._setRunning(false);
        this._showError(err.message);
      }
    },

    async _uploadAllFiles() {
      const roles = ["atc", "clone", "smodilog", "trnspacet"];
      for (const role of roles) {
        const file = this._files[role];
        if (!file) continue;

        const fd = new FormData();
        fd.append("file", file);
        fd.append("role", role);

        // After the first upload, pin all subsequent requests to the same CF instance.
        // X-CF-App-Instance: "<app-guid>:<index>" forces the gorouter to route to that instance.
        const headers = {};
        if (this._cfAppInstance) {
          headers["X-CF-App-Instance"] = this._cfAppInstance;
        }

        const resp = await fetch("/atc/upload", { method: "POST", headers, body: fd });
        const result = await resp.json();
        if (result.error) throw new Error(`Upload failed (${role}): ${result.error}`);

        // Capture instance from first upload so all subsequent requests go there
        if (!this._cfAppInstance && result.cfAppInstance) {
          this._cfAppInstance = result.cfAppInstance;
          // Persist so Results page (different controller) uses the same instance
          try { sessionStorage.setItem("cfAppInstance", result.cfAppInstance); } catch (_) {}
          // Plant __VCAP_ID__ affinity cookie by hitting /atc/pin with X-CF-App-Instance.
          // The gorouter rewrites it to its own affinity value for this instance so the
          // browser sends it on ALL future requests — including <a> download link clicks
          // that cannot carry custom headers.
          try {
            await fetch("/atc/pin", {
              method: "POST",
              headers: { "X-CF-App-Instance": result.cfAppInstance },
            });
          } catch (_) {}
        }

        this._fileIds[role] = result.fileId;

        // Update status label
        const statusMap = {
          atc:       "atcFileStatus",
          clone:     "cloneFileStatus",
          smodilog:  "smodilogStatus",
          trnspacet: "trnspacetStatus",
        };
        this.byId(statusMap[role])
          ?.setText(`${result.filename} (${_formatSize(result.size)})`)
          .setState("Success");
      }
    },

    async _pollJob(jobId) {
      let consecutiveErrors = 0;
      const poll = async () => {
        try {
          const pollHeaders = {};
          if (this._cfAppInstance) pollHeaders["X-CF-App-Instance"] = this._cfAppInstance;
          const resp = await fetch(
            `/api/atc/Jobs(id='${jobId}')?$select=status,statusMsg`,
            { headers: pollHeaders }
          );

          // Non-2xx (e.g. 502 from CF during restart) — retry up to 5 times before giving up
          if (!resp.ok) {
            consecutiveErrors++;
            if (consecutiveErrors >= 5) {
              this._setRunning(false);
              this._showError(`Analysis server unavailable (HTTP ${resp.status}). Please try again.`);
              return;
            }
            setTimeout(poll, 5000);
            return;
          }
          consecutiveErrors = 0;

          const data   = await resp.json();
          const status = (data.value?.[0] || data).status;
          const msg    = (data.value?.[0] || data).statusMsg || "";

          this.byId("statusText").setText(msg);

          if (status === "done") {
            this._setRunning(false);
            MessageToast.show("Analysis complete!", { duration: 2000 });
            this.getOwnerComponent().getRouter().navTo("results", { jobId });
            return;
          }
          if (status === "failed") {
            this._setRunning(false);
            this._showError(`Analysis failed: ${msg}`);
            return;
          }
          setTimeout(poll, 4000);
        } catch (err) {
          // Network error or JSON parse failure — retry a few times before surfacing
          consecutiveErrors++;
          if (consecutiveErrors >= 5) {
            this._setRunning(false);
            this._showError(`Polling error: ${err.message}`);
            return;
          }
          setTimeout(poll, 5000);
        }
      };
      setTimeout(poll, 2000);
    },

    _setRunning(running, msg) {
      const statusBox = this.byId("statusBox");
      const runBtn    = this.byId("runBtn");
      statusBox.setVisible(running);
      if (!running) runBtn.setEnabled(this._isFormValid());
      else runBtn.setEnabled(false);
      if (msg) this.byId("statusText").setText(msg);
    },

    _isFormValid() {
      const customer = this.byId("customerInput")?.getValue().trim() || "";
      const m        = this._analysisMode;
      const hasAtc   = !!this._files.atc;
      const hasClone = !!this._files.clone;
      const hasSmod  = !!this._files.smodilog;
      const hasTrns  = !!this._files.trnspacet;

      if (!customer) return false;
      switch (m) {
        case "tua_only":      return hasSmod && hasTrns;
        case "atc_clone":     return hasAtc && hasClone;
        case "atc_tua":       return hasAtc && hasSmod && hasTrns;
        case "atc_tua_clone": return hasAtc && hasClone && hasSmod && hasTrns;
        default:              return hasAtc; // "atc"
      }
    },

    _updateRunButton() {
      this.byId("runBtn")?.setEnabled(this._isFormValid());
    },

    _showError(msg) {
      const strip = this.byId("errorStrip");
      strip.setText(msg).setVisible(true);
    },

    onCloseError() {
      this.byId("errorStrip").setVisible(false);
    },
  });
});

function _formatSize(bytes) {
  if (!bytes) return "0 B";
  const kb = bytes / 1024;
  return kb < 1024 ? `${kb.toFixed(1)} KB` : `${(kb/1024).toFixed(1)} MB`;
}
