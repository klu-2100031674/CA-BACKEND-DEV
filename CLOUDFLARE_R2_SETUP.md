# Cloudflare R2 Setup Guide

## Step 1: Get R2 Access Keys

1. Go to Cloudflare Dashboard: https://dash.cloudflare.com
2. Select your account
3. Navigate to **R2** in the left sidebar
4. Click **Manage R2 API Tokens**
5. Click **Create API Token**
6. Give it a name: "DPR Excel Storage"
7. Permissions: **Object Read & Write**
8. Click **Create API Token**
9. **IMPORTANT**: Copy both keys immediately (you won't see them again):
   - Access Key ID
   - Secret Access Key

## Step 2: Get Account ID

1. In Cloudflare Dashboard, click on R2
2. Look at the URL or the sidebar - you'll see your Account ID
3. It looks like: `a1b2c3d4e5f6g7h8i9j0k1l2m3n4o5p6`

## Step 3: Update .env File

Update your `.env` file with these values:

```env
# Cloudflare R2 Storage
R2_ENDPOINT=https://<YOUR_ACCOUNT_ID>.r2.cloudflarestorage.com
R2_ACCESS_KEY_ID=<YOUR_ACCESS_KEY_ID>
R2_SECRET_ACCESS_KEY=<YOUR_SECRET_ACCESS_KEY>
R2_BUCKET_NAME=dpr-excel-storage
R2_PUBLIC_URL=https://your-custom-domain.com
```

Replace:
- `<YOUR_ACCOUNT_ID>` - Your Cloudflare Account ID
- `<YOUR_ACCESS_KEY_ID>` - The Access Key ID from Step 1
- `<YOUR_SECRET_ACCESS_KEY>` - The Secret Access Key from Step 1

## Step 4: (Optional) Setup Custom Domain

If you want to use a custom domain instead of the default R2 URL:

1. In R2 bucket settings, click **Settings**
2. Under **Public Access**, click **Connect Domain**
3. Enter your domain (e.g., `files.yourapp.com`)
4. Follow DNS setup instructions
5. Update `R2_PUBLIC_URL` in .env with your custom domain

## Folder Structure

Files will be automatically organized as:
- Excel files: `{userEmail}/excel/{filename}.xlsx`
- PDF files: `{userEmail}/pdf/{filename}.pdf`

Example:
```
dpr-excel-storage/
  └── user@example.com/
      ├── excel/
      │   └── TERM_LOAN_CC_2025-12-14.xlsx
      └── pdf/
          └── TERM_LOAN_CC_Report_2025-12-14.pdf
```

## Testing

After adding credentials to .env:
1. Restart your backend server
2. Generate a new report
3. Files should upload to R2 instead of local storage
4. Admin can download from validation panel
5. Check Cloudflare R2 dashboard to see uploaded files

## Benefits

✅ **10GB free storage** (not just 5GB!)
✅ **No egress fees** (unlimited downloads)
✅ **Global CDN** (fast worldwide access)
✅ **S3-compatible** (easy to migrate)
✅ **Always free** (no 12-month limit)

## Troubleshooting

If uploads fail:
1. Check R2 credentials are correct in .env
2. Verify Account ID in endpoint URL
3. Ensure API token has Read & Write permissions
4. Check backend logs for detailed errors
5. System will fallback to database storage if R2 fails
