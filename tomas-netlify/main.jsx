import React from "react";
import ReactDOM from "react-dom/client";
import "./storage.js"; // define window.storage antes de montar la app
import App from "./App.jsx";
import Auth from "./Auth.jsx";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <Auth>
      <App />
    </Auth>
  </React.StrictMode>
);
