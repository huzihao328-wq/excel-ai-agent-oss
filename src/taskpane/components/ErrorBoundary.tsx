import { Component, ErrorInfo, ReactNode } from "react";
import { setCurrentConversationId } from "../../core/conversations";

interface Props {
  children: ReactNode;
}

interface State {
  error: Error | null;
}

/**
 * 渲染出错时的兜底：没有它，任何一个渲染异常都会让整个任务窗格变成空白，
 * 而且重开插件仍会加载同一段出问题的对话，用户无路可走。
 */
export default class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null };

  static getDerivedStateFromError(error: Error): State {
    return { error };
  }

  componentDidCatch(error: Error, info: ErrorInfo): void {
    console.error("[taskpane] render error", error, info.componentStack);
  }

  private reload = () => {
    window.location.reload();
  };

  private resetConversation = () => {
    try {
      setCurrentConversationId(null);
    } catch {
      /* 存储不可用时也照常重载 */
    }
    window.location.reload();
  };

  render(): ReactNode {
    const { error } = this.state;
    if (!error) return this.props.children;
    return (
      <div className="fatal-panel" role="alert">
        <h2>界面出了点问题</h2>
        <p>插件界面渲染时发生错误。工作簿里的数据不受影响，重新加载即可继续。</p>
        <pre className="fatal-detail">{error.message || String(error)}</pre>
        <div className="fatal-actions">
          <button className="btn primary" onClick={this.reload}>重新加载</button>
          <button className="btn secondary" onClick={this.resetConversation} title="如果每次打开都报错，先从一个新对话开始">
            从新对话重新加载
          </button>
        </div>
      </div>
    );
  }
}
