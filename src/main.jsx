import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";
import Verify from "./Verify.jsx";
import "./index.css";

const path = window.location.pathname;

// GitHub Pages:
// /diplom/verify/0001
const isVerifyPage = path.includes("/verify/");

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    {isVerifyPage ? <Verify /> : <App />}
  </React.StrictMode>
);