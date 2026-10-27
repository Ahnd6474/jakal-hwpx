# Document privacy and security

Uploaded files are private runtime data, not part of this open-source project.
Do not attach real documents to a public issue, pull request, test fixture or
security report. Use a minimal synthetic example and redact names and contents.
Report exploitable issues privately through the repository's private security
advisory channel when enabled; do not publish document URLs or session cookies.

The server uses anonymous, server-issued browser sessions rather than accounts.
Possession of the HttpOnly session cookie grants access to that session's files.
It checks ownership on every document, revision, job, edit and download request.
Sharing a browser profile shares that session until **보안·사용 종료** is used.

For public hosting, follow [the deployment and privacy guide](docs/document-privacy.md).
Production requires a dedicated HTTPS origin, private storage, a restricted
service account and a single API/queue instance. Source code publication does
not publish uploaded data; publishing the runtime volume or real test samples does.

This is not end-to-end encryption. A server administrator, compromised service
account or malicious code running in the same browser origin can still access
data. Use an encrypted disk/volume and restricted backups for protection at rest.
Session separation does not replace OS/container isolation of untrusted parsers.
