import { Component, type ReactNode } from 'react';

/** A missing optional chunk must never take the open document down with it. */
export class ToolBoundary extends Component<{ children: ReactNode; onClose: () => void }, { failed: boolean }> {
  state = { failed: false };
  static getDerivedStateFromError() { return { failed: true }; }
  render() {
    if (!this.state.failed) return this.props.children;
    return <div className="tool-load-failure" role="alert"><p>도구 화면을 불러오지 못했습니다. 편집 내용을 저장한 뒤 새로고침해 주세요.</p><button className="outline-button" onClick={this.props.onClose}>닫고 편집 계속</button></div>;
  }
}
