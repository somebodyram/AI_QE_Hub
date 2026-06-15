// srv/cdata.js
const axios = require("axios");

const CDATA_URL = process.env.CDATA_URL;
const CDATA_AUTH = process.env.CDATA_AUTH;

let sessionId = null;

async function cdataRequest(method, params = {}, reqId = 1) {
    const headers = {
        "Content-Type": "application/json",
        "Accept": "application/json, text/event-stream",
        "Authorization": CDATA_AUTH,
    };

    if (sessionId) {
        headers["Mcp-Session-Id"] = sessionId;
    }

    const body = {
        jsonrpc: "2.0",
        id: reqId,
        method,
        params,
    };

    const response = await axios.post(CDATA_URL, body, {
        headers,
        responseType: "text",      // get raw text so we can handle SSE
        transformResponse: [d => d] // prevent axios from auto-parsing
    });

    // Capture session ID for subsequent calls
    const sid = response.headers["mcp-session-id"];
    if (sid) sessionId = sid;

    const contentType = response.headers["content-type"] || "";
    const results = [];

    if (contentType.includes("text/event-stream")) {
        for (const line of response.data.split("\n")) {
            const trimmed = line.trim();
            if (trimmed.startsWith("data:")) {
                const data = trimmed.slice(5).trim();
                if (data && data !== "[DONE]") {
                    results.push(JSON.parse(data));
                }
            }
        }
    } else {
        results.push(JSON.parse(response.data));
    }

    return results;
}

// Must be called once before any query
async function initCData() {
    console.log("[CData] Initializing session...");
    const result = await cdataRequest("initialize", {
        protocolVersion: "2024-11-05",
        capabilities: {},
        clientInfo: { name: "cap-node-client", version: "1.0" },
    });
    console.log("[CData] Session initialized. ID:", sessionId);
    return result;
}

// Run any SQL query through CData
async function queryData(sql) {
    console.log("[CData] Running query:", sql);
    return cdataRequest("tools/call", {
        name: "queryData",
        arguments: { query: sql },
    });
}

module.exports = { initCData, queryData };

module.exports = { initCData, queryData, cdataRequest };