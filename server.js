import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import {
  testLogin,
  discoverCoursesOnly,
  importAllFromPowerCampus,
  dryRunPushAttendance,
  submitPushAttendance,
} from "./bot.js";

dotenv.config();

const app = express();
app.use(cors());
app.use(express.json({ limit: "50mb" }));

const PORT = process.env.PORT || 8787;

function requireCredentials(req, res) {
  const { username, password } = req.body || {};
  if (!username || !password) {
    res.status(400).json({ ok: false, error: "username and password are required" });
    return null;
  }
  return { username, password };
}

app.get("/health", (req, res) => {
  res.json({
    ok: true,
    service: "powercampus-bot-server",
    version: "v8-push-dry-run-submit",
    submitUnlocked: process.env.ALLOW_POWER_CAMPUS_SUBMIT === "true",
  });
});

app.post("/test-login", async (req, res) => {
  try {
    const credentials = requireCredentials(req, res);
    if (!credentials) return;
    const result = await testLogin(credentials);
    res.json({ ok: true, result });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message, stack: process.env.NODE_ENV === "development" ? err.stack : undefined });
  }
});

app.post("/discover-courses", async (req, res) => {
  try {
    const credentials = requireCredentials(req, res);
    if (!credentials) return;
    const result = await discoverCoursesOnly(credentials);
    res.json({ ok: true, ...result });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message, stack: process.env.NODE_ENV === "development" ? err.stack : undefined });
  }
});

app.post("/import-all", async (req, res) => {
  try {
    const credentials = requireCredentials(req, res);
    if (!credentials) return;
    const result = await importAllFromPowerCampus(credentials);
    res.json(result);
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message, stack: process.env.NODE_ENV === "development" ? err.stack : undefined });
  }
});

app.post("/dry-run-push", async (req, res) => {
  try {
    const credentials = requireCredentials(req, res);
    if (!credentials) return;
    const { attendanceUrl, records, localSummary } = req.body || {};
    if (!attendanceUrl) return res.status(400).json({ ok: false, error: "attendanceUrl is required" });
    if (!Array.isArray(records)) return res.status(400).json({ ok: false, error: "records array is required" });
    const result = await dryRunPushAttendance({ ...credentials, attendanceUrl, records, localSummary });
    res.json(result);
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message, stack: process.env.NODE_ENV === "development" ? err.stack : undefined });
  }
});

app.post("/submit-push", async (req, res) => {
  try {
    const credentials = requireCredentials(req, res);
    if (!credentials) return;
    const { attendanceUrl, records, localSummary } = req.body || {};
    if (!attendanceUrl) return res.status(400).json({ ok: false, error: "attendanceUrl is required" });
    if (!Array.isArray(records)) return res.status(400).json({ ok: false, error: "records array is required" });
    const result = await submitPushAttendance({ ...credentials, attendanceUrl, records, localSummary });
    res.json(result);
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message, stack: process.env.NODE_ENV === "development" ? err.stack : undefined });
  }
});

app.listen(PORT, () => {
  console.log(`[OK] PowerCampus bot server running on http://localhost:${PORT}`);
});
