const express = require("express");
const path = require("path");
const cors = require("cors");
const app = express();
const port = process.env.PORT || 3000;

app.use(cors());

// Allow content to load in Outlook iframe
app.use((req, res, next) => {
  res.setHeader('X-Frame-Options', 'ALLOWALL');
  res.setHeader('Content-Security-Policy', "frame-ancestors *;");
  next();
});

// Set proper MIME types for JavaScript files
app.use(express.static(path.join(__dirname), {
  setHeaders: (res, path) => {
    if (path.endsWith('.js')) {
      res.setHeader('Content-Type', 'application/javascript');
    }
  }
}));

// Serve index.html at root
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

// Start HTTP server (Render provides HTTPS automatically)
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
