import { Component, type ReactNode } from 'react';

export class ApplicationBoundary extends Component<{ children: ReactNode }, { failed: boolean }> {
  state = { failed: false };
  static getDerivedStateFromError() { return { failed: true }; }
  render() {
    if (!this.state.failed) return this.props.children;
    return <main className="application-failure" role="alert"><h1>화면을 표시하지 못했습니다.</h1>
      <p>서버에 저장된 문서와 브라우저의 임시본은 삭제하지 않았습니다. 다시 열면 저장된 임시본을 확인할 수 있습니다. 마지막 입력은 임시보관되지 않았을 수 있습니다.</p>
      <button className="primary-button" onClick={() => window.location.reload()}>페이지 다시 불러오기</button>
    </main>;
  }
}
