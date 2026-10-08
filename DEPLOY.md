# Deploying to the cloud (Fly.io)

The app keeps its database and Excel files on disk, so it needs a host with a persistent volume
(not Vercel/serverless). Run **exactly one** instance.

```sh
brew install flyctl && fly auth login
fly launch --no-deploy --copy-config --name <unique-app-name>   # keep the existing fly.toml
fly volumes create atm_data --size 1 --region iad
fly secrets set ADMIN_PASSWORD='<choose>' REP_PASSWORD='<choose>'   # used on first start only
fly deploy --ha=false
```

Then open `https://<unique-app-name>.fly.dev`, log in as Admin / Treasurer, and upload the preorder sheet on the Lunch Check-in screen.

- Download the Excel files from Day Report regularly as a backup.
- Any Docker host works (Railway, Render, a VPS): mount a volume at `/data` and set `DATA_DIR=/data`, `TRUST_PROXY=1`.
