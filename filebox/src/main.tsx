import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import { installContextMenuBlocker } from "./disableContextMenu";
import "./styles.css";

installContextMenuBlocker();

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
