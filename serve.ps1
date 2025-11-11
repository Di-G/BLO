param(
	[int]$Port = 5500
)

Write-Host "Starting local server on http://localhost:$Port ..."
Set-Location -LiteralPath $PSScriptRoot

function Start-PythonServer {
	if (Get-Command py -ErrorAction SilentlyContinue) {
		py -m http.server $Port
		return $true
	}
	if (Get-Command python -ErrorAction SilentlyContinue) {
		python -m http.server $Port
		return $true
	}
	return $false
}

function Start-NodeServer {
	if (Get-Command npx -ErrorAction SilentlyContinue) {
		npx --yes serve -l $Port .
		return $true
	}
	if (Get-Command node -ErrorAction SilentlyContinue) {
		# Minimal static file server in Node if 'serve' isn't available
		$serverJs = @"
const http = require('http');
const fs = require('fs');
const path = require('path');
const port = process.env.PORT || $Port;
const base = process.cwd();
const mime = {
	'.html': 'text/html; charset=utf-8',
	'.css': 'text/css; charset=utf-8',
	'.js': 'application/javascript; charset=utf-8',
	'.json': 'application/json; charset=utf-8',
	'.png': 'image/png',
	'.jpg': 'image/jpeg',
	'.jpeg': 'image/jpeg',
	'.svg': 'image/svg+xml; charset=utf-8',
	'.ico': 'image/x-icon'
};
http.createServer((req, res) => {
	let reqPath = decodeURI(req.url.split('?')[0]);
	if (reqPath === '/') reqPath = '/index.html';
	const filePath = path.join(base, reqPath);
	fs.stat(filePath, (err, stat) => {
		if (err || !stat.isFile()) {
			res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
			return res.end('Not Found');
		}
		const ext = path.extname(filePath).toLowerCase();
		res.writeHead(200, { 'Content-Type': mime[ext] || 'application/octet-stream' });
		fs.createReadStream(filePath).pipe(res);
	});
}).listen(port, () => {
	console.log('Serving', base, 'on http://localhost:' + port);
});
"@
		$tmp = Join-Path $env:TEMP "simple_static_server.js"
		$serverJs | Set-Content -LiteralPath $tmp -Encoding UTF8
		$env:PORT = "$Port"
		node $tmp
		return $true
	}
	return $false
}

if (Start-PythonServer) { exit 0 }
if (Start-NodeServer)   { exit 0 }

Write-Warning "No suitable runtime found."
Write-Host "Try one of the following and re-run serve.ps1:"
Write-Host " - Install Python:  https://www.python.org/downloads/  (then run: py -m http.server 5500)"
Write-Host " - Install Node.js: https://nodejs.org/  (then run: npx serve -l 5500 .)"
Read-Host "Press Enter to close"


