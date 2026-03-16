const http = require("http");
const fs = require("fs");
const os = require("os");
const path = require("path");

const port = Number(process.env.PORT || 4173);
const host = process.env.HOST || "0.0.0.0";
const root = __dirname;

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".webmanifest": "application/manifest+json; charset=utf-8",
};

const server = http.createServer((request, response) => {
  const rawPath = request.url === "/" ? "/index.html" : request.url || "/index.html";
  const safePath = path.normalize(decodeURIComponent(rawPath)).replace(/^(\.\.[/\\])+/, "");
  const filePath = path.join(root, safePath);

  if (!filePath.startsWith(root)) {
    response.writeHead(403);
    response.end("Forbidden");
    return;
  }

  fs.readFile(filePath, (error, buffer) => {
    if (error) {
      response.writeHead(404);
      response.end("Not Found");
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const type = mimeTypes[ext] || "application/octet-stream";
    response.writeHead(200, {
      "Content-Type": type,
      "Cache-Control": "no-cache",
    });
    response.end(buffer);
  });
});

server.listen(port, host, () => {
  const urls = getAccessUrls(port, host);
  console.log("System English Words is available at:");
  urls.forEach((url) => console.log(`  ${url}`));
});

function getAccessUrls(portNumber, listenHost) {
  const urls = [];
  if (listenHost === "0.0.0.0") {
    urls.push(`http://127.0.0.1:${portNumber}`);

    const interfaces = os.networkInterfaces();
    Object.values(interfaces)
      .flat()
      .filter(Boolean)
      .filter((info) => info.family === "IPv4" && !info.internal)
      .forEach((info) => {
        urls.push(`http://${info.address}:${portNumber}`);
      });
    return [...new Set(urls)];
  }

  return [`http://${listenHost}:${portNumber}`];
}
