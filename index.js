const express = require('express');
const https = require('https');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;

// Load dev certs (from office-addin-dev-certs)
const cert = fs.readFileSync(path.join(process.env.HOME || process.env.USERPROFILE, '.office-addin-dev-certs', 'localhost.crt'));
const key = fs.readFileSync(path.join(process.env.HOME || process.env.USERPROFILE, '.office-addin-dev-certs', 'localhost.key'));

// Allow your content to load in Outlook iframe
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('X-Frame-Options', 'ALLOWALL');
  next();
});

// Serve everything from /public folder
app.use(express.static(path.join(__dirname, '../public')));

// Start HTTPS server
https.createServer({ key, cert }, app).listen(PORT, () => {
  console.log(`Add-in is available at https://localhost:${PORT}`);
});
