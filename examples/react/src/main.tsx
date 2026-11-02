import React from "react";
import { createRoot } from "react-dom/client";
import "@dukeofwest/universal-file-preview/styles.css";
import App from "./App.js";
import "./app.css";

const rootElement = document.getElementById("root");
if (!rootElement) throw new Error("React root element is missing.");

createRoot(rootElement).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
