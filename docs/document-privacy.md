# 문서 비공개 저장과 공개 배포

목표는 다른 사이트 사용자에게 원본·수정본·본문·표·검사 기록·작업 결과가 노출되지 않도록 하는 것이다. 소스 공개와 실제 파일 공개는 별개다. 실제 사용자 문서를 예제나 공개 이슈에 첨부하지 않는다.

## 적용한 경계

- 서버가 생성한 256비트 세션만 사용한다. SQLite에는 쿠키 원문 대신 SHA-256 해시, 비밀이 아닌 브라우저 보관함 구분자, 만료·폐기 상태를 기록한다. 임의의 64자리 문자열은 서버 세션으로 채택하지 않는다.
- 원본·수정본·텍스트·검사 기록·구역 설정·과거 버전·작업 조회 및 모든 변경 API에서 소유자와 문서 만료를 확인한다. 다른 세션의 문서 ID·작업 ID·다운로드 주소를 알아도 404다. Range 요청도 우회하지 못한다.
- 보관함 구분자는 인증 수단이 아니다. 새 UI는 `X-Jakal-Workspace`를 함께 보내며 쿠키의 실제 세션과 다르면 본문 처리 전에 401을 반환한다. 세션이 바뀐 탭의 업로드·저장이 다른 세션으로 들어가는 것을 막는다.
- 세션 확인을 끝내기 전에는 문서 UI를 표시하지 않는다. 포커스 복귀·표시 상태 변경·30초 주기 조회와 동일 출처 탭 간 종료 알림으로 만료/종료된 화면을 닫는다. 이미 완료된 다운로드나 화면 촬영을 회수하는 기능은 아니다.
- 원문 편집 IndexedDB, 배경 양식 IndexedDB, 양식 설정·최근 문서·텍스트 임시본 키를 세션별로 구분한다. 출처가 같다는 이유로 다른 세션의 보관함을 기본 표시하지 않는다. 이 구분은 같은 기기에서 개발자 도구를 쓸 수 있는 사람이나 XSS에 대한 암호화 경계는 아니다.
- 사이드바 **보안·사용 종료**는 명시적 삭제 동의를 받은 뒤 세션을 폐기하고 문서를 즉시 만료시킨다. 복사한 쿠키도 다시 쓸 수 없다. 파일은 워커가 처리 중인 작업과 파일 잠금이 풀린 뒤 정리한다. 통상 60초 주기이고 서버가 꺼져 있으면 다음 실행까지 디스크에 남는다. 브라우저의 임시본·내 양식도 지우며 HTTPS에서는 `Clear-Site-Data`를 함께 보낸다. 이 기능 때문에 다른 앱과 저장 공간을 공유하지 않는 전용 origin을 사용한다.
- 쿠키는 HttpOnly·SameSite=Strict·Path=/다. HTTPS에서는 Domain 속성 없는 `__Host-jakal_session; Secure`를 사용한다. 운영 모드는 HTTPS origin 일치·HSTS를 강제하고 API 문서를 공개하지 않는다.
- Host 및 Origin을 검사하고 cross-site/same-site 브라우저의 API 요청도 거절한다. 개발 모드는 루프백 접속만 허용한다. 프록시를 신뢰하는 IP는 명시적으로 제한한다.
- 비공개 API는 `no-store`, `Vary: Cookie`, `Pragma: no-cache`, `Expires: 0`, CORP=same-origin, Referrer-Policy=no-referrer를 사용한다. 문서는 정적 경로로 제공하지 않으며 데이터 경로와 공개 디렉터리가 겹치면 시작을 거부한다.
- 정적 서버도 숨김 경로·업로드 디렉터리·DB·개인키 및 일반 HWP/HWPX/PDF를 거절한다. 소개용 공개 샘플 두 개만 파일 해시가 검토한 값과 같을 때 제공한다. 실수로 공개 폴더에 복사된 문서도 일반 정적 다운로드로 노출되지 않게 하는 추가 방어다.
- POSIX 저장 디렉터리는 0700, DB는 0600이며 새 파일은 umask 077을 적용한다. Windows 운영 모드는 현재 서비스 계정·SYSTEM·Administrators만 허용하는 ACL을 적용한다. `scripts/secure_private_storage.ps1`로 로컬 저장소에도 같은 설정을 적용할 수 있다. ACL 적용에 실패하면 운영 시작을 중단한다.
- 파일 해석 프로세스에는 필요한 OS/런타임 환경만 전달한다. API 키·DB 접속 문자열·NODE_OPTIONS 같은 호스트 설정을 그대로 물려주지 않는다. ZIP/XML 검증, DTD/외부 엔티티 차단, 업로드·압축 해제·메모리·시간 제한은 유지한다. 이는 별도 OS 샌드박스나 백신을 대신하지 않는다.
- 일반 접근 로그에서 문서 URL을 남기지 않도록 기본 실행의 Uvicorn 접근 로그를 끈다. 프록시·호스팅 제공자의 로그 정책도 별도로 설정해야 한다.

익명 세션은 마지막 사용 이후 기본 24시간, 최대 7일 동안 유효하다. 파일 보관 기간은 각 문서 업로드 이후 24시간이다. 쿠키가 만료/폐기되어도 새 세션으로 파일 권한을 승계하지 않는다. 로컬 업그레이드에 한해 시작 시 이미 보관 중인 문서 소유자의 기존 쿠키를 한 번 채택한다. 운영 모드는 이 호환 경로를 사용하지 않는다.

기존 버전의 브라우저 임시본·내 양식은 소유자 정보가 없으므로 새 보관함으로 일괄 승계하지 않는다. 해당 데이터를 다른 세션에 노출하지 않기 위한 조치다. 이전 버전에서 내보낸 양식 JSON은 가져올 수 있다. 명시적인 사용 종료는 이전 브라우저 저장소도 정리한다.

## 공개 서버 실행

전용 도메인 예: `https://editor.example.com`. Python은 한 프로세스로 실행하고 외부 접속은 HTTPS 역방향 프록시를 통해서만 받는다. 다음은 환경 변수의 예이며 애플리케이션이 `.env`를 자동으로 읽지는 않는다.

```bash
export JAKAL_ENV=production
export JAKAL_PUBLIC_ORIGIN=https://editor.example.com
export JAKAL_DATA_DIR=/var/lib/jakal-private
export JAKAL_HOST=127.0.0.1
export JAKAL_PORT=8000
export JAKAL_TRUSTED_PROXIES=127.0.0.1
python -m server
```

Windows에서는 동일한 값을 `$env:JAKAL_ENV='production'` 형태로 지정한다. `JAKAL_DATA_DIR`는 서비스 계정 전용 디렉터리로 지정한다. 운영 시작 시 ACL 스크립트가 실행되므로 서명된 스크립트를 요구하는 환경은 해당 정책에 맞게 배포한다. 운영 설정 없이 외부 HTTP로 제공하는 방식은 지원하지 않는다.

프록시는 애플리케이션의 공개 도메인 Host와 실제 HTTPS scheme을 전달해야 한다. TLS는 프록시에서 종료하고 Python 포트는 방화벽으로 외부 접근을 막는다. **저장소 전체·`.web-data`·사용자 데이터 볼륨을 nginx/IIS의 정적 root나 alias로 연결하면 안 된다.** 프록시 설정 예:

```nginx
location / {
    proxy_pass http://127.0.0.1:8000;
    proxy_set_header Host $host;
    proxy_set_header X-Forwarded-Proto $scheme;
    proxy_set_header X-Forwarded-For $remote_addr;
    proxy_cache off;
    client_max_body_size 21m;
    access_log off;
}
```

HTTPS 인증서 설정, 프록시의 요청 속도 제한, 전체 저장 용량 제한, 전용 서비스 계정·컨테이너, 컨테이너/워커 외부 네트워크 차단, 디스크/백업 암호화는 실제 배포 환경에서 구성한다. `JAKAL_TRUSTED_PROXIES=*`는 운영에서 거부한다. 프록시가 파일 응답을 별도 디스크에 버퍼링하는 환경은 해당 임시 경로와 보관/삭제 정책도 제한한다.

현재 파일·본문·SQLite 작업 데이터는 서버 디스크에서 평문으로 처리한다. 서버 관리자나 서비스 계정이 탈취되었을 때까지 막는 종단간 암호화, 회원 로그인·MFA, 멀티테넌트 OS 격리는 포함하지 않는다. 계정 단위의 장기 보관이 필요하면 검증된 인증 제공자를 연결하고 이 서버의 소유권 경계를 유지해야 한다.

## 소스 공개 전 확인

`python scripts/check_private_artifacts.py`를 실행한다. CI도 동일 검사를 실행한다. Git 추적 파일 중 업로드 저장소·SQLite·실제 환경 변수 파일·개인키·문서 스냅숏·작업 임시 파일·공개 웹 경로의 HWP/HWPX/PDF가 있으면 실패한다. 예외는 검토한 생성형 소개용 샘플 두 개의 정확한 해시뿐이다. 내용이 바뀌면 다시 실패한다. 실제 사용자 문서를 샘플 이름으로 바꾸어도 통과하지 못한다.

이 검사는 문서 전체의 개인정보 탐지나 Git 과거 이력 검사가 아니다. 기존 `examples/`, `debug sample/` 등 추적 중인 문서는 별도로 권리·개인정보를 검토해야 한다. 과거에 비공개 파일을 커밋했다면 `.gitignore`만으로 이력에서 제거되지 않는다. 공개 전 이력 정리와 노출된 인증 정보 폐기가 필요하다.

설계 참고: [OWASP 파일 업로드 지침](https://cheatsheetseries.owasp.org/cheatsheets/File_Upload_Cheat_Sheet.html), [세션 관리 지침](https://cheatsheetseries.owasp.org/cheatsheets/Session_Management_Cheat_Sheet.html), [접근 권한 검사 지침](https://cheatsheetseries.owasp.org/cheatsheets/Authorization_Cheat_Sheet.html).
