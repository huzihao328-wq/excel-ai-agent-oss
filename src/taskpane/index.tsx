/* global Office */
import { createRoot } from "react-dom/client";
import { hydrateUserData } from "../core/durableStore";
import App from "./App";
import ErrorBoundary from "./components/ErrorBoundary";
import "./styles.css";
import "./workspace.css";
import "./motion.css";
import { installMotionTokens } from "./motion";

installMotionTokens();

Office.onReady(() => {
  void hydrateUserData().catch(() => undefined).then(() => {
    const container = document.getElementById("root");
    if (container) {
      createRoot(container).render(
        <ErrorBoundary>
          <App />
        </ErrorBoundary>
      );
    }
  });
});
