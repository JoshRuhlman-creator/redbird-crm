// Loads /opt/redbird-sync/.env (KEY=value lines) into process.env.
// Keeps the Cultivera login out of the code. The file is chmod 600.
const fs = require('fs');
const path = require('path');
const file = path.join(__dirname, '.env');
if (fs.existsSync(file)) {
  fs.readFileSync(file, 'utf8').split(/\r?\n/).forEach(line => {
    const m = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*)\s*$/);
    if (m && process.env[m[1]] === undefined) process.env[m[1]] = m[2].replace(/^(['"])(.*)\1$/, '$2');
  });
}
