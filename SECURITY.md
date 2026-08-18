# Security policy

Please report vulnerabilities privately through GitHub's security advisory feature rather than a public issue.

The app processes untrusted documents locally. Changes to archive extraction, legacy-office conversion, path handling, generated-directory cleanup, external commands, or HTML reports require tests that cover traversal and command-injection risks.

Never include source documents, API keys, private AI prompts, or local model weights in a report.
