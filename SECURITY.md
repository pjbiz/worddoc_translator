# Security Policy

## Overview

This document outlines security considerations for the Word Document Translator tool.

## Reporting Security Vulnerabilities

If you discover a security vulnerability, please report it responsibly:

1. **Do NOT** create a public GitHub issue for security vulnerabilities
2. Email the maintainer directly with details of the vulnerability
3. Include steps to reproduce the issue
4. Allow reasonable time for a fix before public disclosure

## Security Measures Implemented

### API Key Protection
- **Web UI**: API keys are entered by users and stored only in browser session state
- **Session-only storage**: Keys are never saved to disk or persistent storage
- **Validation**: Keys are validated before use with a minimal API call
- **Masked display**: Only first 8 and last 4 characters are shown in the UI
- **CLI**: Uses environment variables or `.env` files (excluded from git)
- API keys are never logged or displayed in error messages

### Input Validation
- File uploads are restricted to `.docx` format only
- Maximum file size limit: 10 MB per file
- Maximum files per upload: 20 files
- Text content is truncated at 50,000 characters per block
- File paths are validated to prevent path traversal attacks

### Error Handling
- Internal error details are logged but not exposed to users
- Generic error messages are shown to prevent information leakage

### Dependencies
- All dependencies are pinned to specific versions in `requirements.txt`
- Regular dependency updates are recommended

## Security Best Practices for Users

### API Key Management
1. **Web UI**: Your key is stored only in your browser session - it's cleared when you close the tab
2. **CLI**: Never commit your `.env` file to version control
3. Use a dedicated API key for this application
4. Set usage limits on your OpenAI API key at [OpenAI Platform](https://platform.openai.com/account/limits)
5. Rotate your API key periodically
6. Monitor your API usage for unexpected activity

### Deployment Considerations
1. **Local Use Only**: This tool is designed for local/personal use
2. **No Authentication**: The web UI has no built-in authentication
3. **Public Deployment**: If deploying publicly, add authentication (e.g., Streamlit's native auth, or a reverse proxy)

### If Deploying Publicly
If you must deploy this publicly, consider:
- Adding authentication via Streamlit Cloud or a reverse proxy
- Implementing rate limiting at the infrastructure level
- Using a dedicated OpenAI API key with strict usage limits
- Monitoring API usage for anomalies
- Running behind HTTPS

## Known Limitations

### Prompt Injection
Document content is sent to the OpenAI API. Malicious documents could potentially contain text designed to manipulate the AI's behavior. The system prompt instructs the model to only translate, but this is not a guaranteed security boundary.

### Data Privacy
- Document content is sent to OpenAI's API for translation
- Review OpenAI's data usage policy for your use case
- Do not use this tool for highly sensitive or confidential documents without appropriate review

### No Rate Limiting
The application does not implement application-level rate limiting. Rely on:
- OpenAI API key usage limits
- Infrastructure-level rate limiting for public deployments

## Dependency Security

Regularly update dependencies to patch security vulnerabilities:

```bash
pip install --upgrade -r requirements.txt
```

Check for known vulnerabilities:

```bash
pip install safety
safety check -r requirements.txt
```

## Version History

| Version | Date | Security Updates |
|---------|------|------------------|
| 1.0.0 | 2026-01 | Initial release with security measures |

## Contact

For security concerns, contact the repository maintainer.

