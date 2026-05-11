# Seafarer ZIP Extractor

Auto-extracts ZIP attachments dropped into `/Email attachments/**/Original/`
in OneDrive, used by the Panama Embassy Seoul seafarer document workflow.

## How it works

A scheduled GitHub Actions workflow (`*/5 * * * *`) authenticates to Microsoft
Graph via an Azure AD app registration (client credentials flow), walks the
OneDrive folder tree, and processes any `.zip` files it finds in `Original/`
subfolders. After successful extraction the source zip is renamed to
`_processed_*.zip` so it isn't picked up again.

## Required GitHub Secrets

- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET`
- `TARGET_USER_UPN`

## Safety guards (hardcoded)

- Refuses to touch anything outside `/Email attachments/`
- Max 500 MB extracted size per zip
- Max 200 files per zip
- Only extracts: PDF, JPG, JPEG, PNG, TIF, TIFF, GIF, BMP
- Skips password-protected zips (renamed `_REJECTED_password_protected_*`)
- Skips corrupt zips (renamed `_REJECTED_corrupt_*`)
- Max 2 levels of nested zips

## Manual run

GitHub Actions tab → "Seafarer ZIP Extractor" → "Run workflow".
