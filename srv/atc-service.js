"use strict";
const cds = require("@sap/cds");
const atcProcessor = require("../lib/atcProcessor");

module.exports = class ATCService extends cds.ApplicationService {

  async init() {
    const { Jobs } = this.entities;

    // ── OData action: runAnalysis ────────────────────────────────────────────
    this.on("runAnalysis", Jobs, async (req) => {
      const jobId = req.params[0].id;

      const [job] = await SELECT.from(Jobs).where({ id: jobId });
      if (!job) return req.error(404, `Job ${jobId} not found`);

      await UPDATE(Jobs).set({ status: "running", statusMsg: "ATC analysis starting..." }).where({ id: jobId });

      // Fire async — do not await so OData response returns immediately
      // analysisMode is stored on the job; useTua is derived inside processJob
      setImmediate(() => atcProcessor.processJob(jobId).catch(console.error));

      return SELECT.one.from(Jobs).where({ id: jobId });
    });

    // ── OData function: getChartData ─────────────────────────────────────────
    this.on("getChartData", Jobs, async (req) => {
      const [job] = await SELECT.from("atc.Jobs").columns("chartData").where({ id: req.params[0].id });
      return job ? (job.chartData || "null") : "null";
    });

    await super.init();
  }
};
