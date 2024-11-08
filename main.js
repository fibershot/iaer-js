// Main script
import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import settings from "./js/appSettings.js";
import { fetchData, initializeGraph } from "./js/graph.js";

// Define __dirname manually
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const PORT = 4070;

// Express configuration
const app = express();
app.use(express.json());
app.use(express.static("public"));
app.use("/css", express.static(path.join(__dirname, "public", "css"), { setHeaders: (res) => res.set("Content-Type", "text/css") }));
app.use("/js", express.static(path.join(__dirname, "public", "js"), { setHeaders: (res) => res.set("Content-Type", "application/javascript") }));
app.use("/fonts", express.static(path.join(__dirname, "public", "fonts"), { setHeaders: (res) => res.set("Content-Type", "font/ttf") }));

initializeGraph(settings);

// Data fetching from graph.js
app.post("/api/fetch-data", async (req, res) => {
    const { serials } = req.body;
    try {
        const results = await fetchData(serials);
        if (results) {
            res.status(200).json({ results });
        } else {
            res.status(403).json({ error: "Nothing returned."});
        }
    } catch (error) {
        console.log(error);
        res.status(500).json({ error: error})
    }
});

// Load HTML page
app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "index.html"));
});

// Open server for port 4070
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});