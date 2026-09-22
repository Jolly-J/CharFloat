import React, { Component, ErrorInfo, ReactNode } from "react";
import ReactDOM from "react-dom/client";
import App from "./App.js";
import { PermissionFloat } from "./components/PermissionFloat.js";
import "./index.css";

interface Props {
  children: ReactNode;
}

interface State {
  hasError: boolean;
  error?: Error;
}

class ErrorBoundary extends Component<Props, State> {
  public state: State = {
    hasError: false
  };

  public static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error };
  }

  public componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error("Uncaught error:", error, errorInfo);
  }

  public render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: 40, backgroundColor: "#0f172a", color: "#f8fafc", minHeight: "100vh", fontFamily: "system-ui" }}>
          <h2 style={{ color: "#ef4444", fontSize: 22, marginBottom: 16 }}>客户端初始化检测到异常</h2>
          <p style={{ color: "#94a3b8", marginBottom: 20 }}>错误原因: {this.state.error?.message}</p>
          <pre style={{ backgroundColor: "#1e293b", padding: 16, borderRadius: 8, overflowX: "auto", fontSize: 12, color: "#cbd5e1" }}>
            {this.state.error?.stack}
          </pre>
          <button
            onClick={() => window.location.reload()}
            style={{ marginTop: 24, padding: "10px 20px", backgroundColor: "#3b82f6", color: "white", border: "none", borderRadius: 6, cursor: "pointer" }}
          >
            重新加载
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <ErrorBoundary>
      {window.location.hash === "#permission-guide" ? <PermissionFloat /> : <App />}
    </ErrorBoundary>
  </React.StrictMode>
);
