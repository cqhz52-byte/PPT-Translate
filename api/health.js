function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.statusCode = 200;
  res.end(JSON.stringify({
    ok: true,
    runtime: "nodejs",
    node: process.version,
    hasFetch: typeof fetch === "function",
    method: req.method,
  }));
}

module.exports = handler;
module.exports.default = handler;
