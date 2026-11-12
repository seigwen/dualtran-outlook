#!/usr/bin/env node
/**
 * print-payload.mjs — `npm run conformance` builds the payload from the real
 * controls (scripts/conformance/payload.mjs → bootPane) and prints the JSON on
 * stdout. Kept separate so tests can import buildPayload() without spawning.
 */
import { buildPayload } from "./payload.mjs";

const payload = await buildPayload();
process.stdout.write(JSON.stringify(payload));
