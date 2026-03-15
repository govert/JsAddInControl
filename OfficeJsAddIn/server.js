const fs = require("fs");
const https = require("https");
const path = require("path");

const rootDir = __dirname;
const port = 3000;
const certificatePath = path.join(rootDir, "localhost-devcert.pfx");
const certificatePassphrase = "testingcertpass";

const contentTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".xml": "application/xml; charset=utf-8"
};

function resolvePath(requestPath) {
  const cleanPath = requestPath === "/" ? "/src/taskpane.html" : requestPath;
  const localPath = path.normalize(path.join(rootDir, cleanPath));

  if (!localPath.startsWith(rootDir)) {
    return null;
  }

  return localPath;
}

async function start() {
  const serverOptions = {
    passphrase: certificatePassphrase,
    pfx: fs.readFileSync(certificatePath)
  };

  const server = https.createServer(serverOptions, (req, res) => {
    const localPath = resolvePath(req.url.split("?")[0]);

    if (!localPath) {
      res.writeHead(400);
      res.end("Bad request");
      return;
    }

    fs.readFile(localPath, (error, fileBuffer) => {
      if (error) {
        res.writeHead(error.code === "ENOENT" ? 404 : 500);
        res.end(error.code === "ENOENT" ? "Not found" : "Server error");
        return;
      }

      const extension = path.extname(localPath).toLowerCase();
      const contentType = contentTypes[extension] || "application/octet-stream";
      res.writeHead(200, {
        "Access-Control-Allow-Origin": "*",
        "Cache-Control": "no-store",
        "Content-Type": contentType
      });
      res.end(fileBuffer);
    });
  });

  server.listen(port, () => {
    console.log(`Office add-in assets available at https://localhost:${port}`);
  });
}

start().catch((error) => {
  console.error(error);
  process.exit(1);
});
