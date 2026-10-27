"""Shared response policy, including requests rejected before a session exists."""


def response_headers(path: str) -> dict[str, str]:
    studio = path.startswith('/studio/')
    headers = {
        'X-Content-Type-Options': 'nosniff',
        'Referrer-Policy': 'no-referrer',
        'Cross-Origin-Resource-Policy': 'same-origin',
        'X-Permitted-Cross-Domain-Policies': 'none',
        'Permissions-Policy': 'camera=(), microphone=(), geolocation=()',
        'Content-Security-Policy': (
            "default-src 'self'; script-src 'self' 'wasm-unsafe-eval'; worker-src 'self' blob:; style-src 'self' 'unsafe-inline'; font-src 'self' data:; img-src 'self' data: blob:; connect-src 'self'; frame-src 'self' blob:; object-src 'none'; frame-ancestors 'self'; base-uri 'self'; form-action 'self'"
            if studio else
            "default-src 'self'; script-src 'self' 'wasm-unsafe-eval'; worker-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; connect-src 'self'; object-src 'none'; frame-ancestors 'none'; base-uri 'self'; form-action 'self'"
        ),
    }
    if path == '/api' or path.startswith('/api/') or path in {'/mcp', '/mcp/'}:
        headers['Cache-Control'] = 'no-store'
        headers['Pragma'] = 'no-cache'
        headers['Expires'] = '0'
        headers['Vary'] = 'Authorization' if path in {'/mcp', '/mcp/'} or path.startswith('/api/connected/') else 'Cookie'
    return headers
