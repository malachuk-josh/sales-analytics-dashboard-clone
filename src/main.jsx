import React from "react";
import { createRoot } from "react-dom/client";
import SalesAnalyticsDashboardApp from "./SalesAnalyticsDashboardApp.jsx";
import "./styles.css";

createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <SalesAnalyticsDashboardApp />
  </React.StrictMode>
);
