# Spotify Audiobook Chart Tracker

Weekly snapshot of Spotify's editorial audiobook playlists into a tracked spreadsheet + per-chart PDFs, with rank-change deltas vs the previous refresh.

## Quickstart: fork this and run your own (5 minutes)

1. Click **Use this template** (top right of this page) → create your private copy under your account.
2. Get a free Gemini API key at https://aistudio.google.com/app/apikey. Free tier covers ~250 lookups/day; with caching, that lasts months. Optional: enable billing for unlimited.
3. Get a free Firecrawl key at https://www.firecrawl.dev/ (1,000 page scrapes/month free). Weekly refresh of 12 playlists uses ~52 credits/month.
4. In your forked repo: **Settings → Secrets and variables → Actions → New repository secret**. Add two:
   - `GEMINI_API_KEY` (the `AIzaSy...` string)
   - `FIRECRAWL_API_KEY` (the `fc-...` string)
5. **Actions** tab → if prompted, click "I understand my workflows, enable them".
6. Edit `playlists.txt` if you want different playlists. Format: `<playlist_id><tab><tab name>`. Find IDs in any Spotify playlist URL: `open.spotify.com/playlist/<this part>`.
7. **Actions → Weekly audiobook refresh → Run workflow** to fire your first run. After ~7 minutes you'll have an updated `audiobooks.xlsx` and 12 PDFs in `output/pdfs/`, committed by the bot.

After step 7 the weekly cron (Monday 13:00 UTC) takes over. Every refresh commits the new xlsx and dated PDFs back to your repo, so the git history doubles as your weekly archive.

## What's tracked

12 playlists from `playlists.txt`:

| Tab | Playlist URL |
|---|---|
| Top | https://open.spotify.com/playlist/37i9dQZF1CHNJaMQB8FMMn |
| Romance | https://open.spotify.com/playlist/37i9dQZF1CHT0u8cbeDlkr |
| Mystery & Thriller | https://open.spotify.com/playlist/37i9dQZF1CHSUrTnfwXo9Y |
| Kids & Family | https://open.spotify.com/playlist/37i9dQZF1CHKw7Rdk4reur |
| Parenting & Relationships | https://open.spotify.com/playlist/37i9dQZF1CHW1rLDWEkz4P |
| History | https://open.spotify.com/playlist/37i9dQZF1CHJqhO7Tb0Aq5 |
| Religion & Spirituality | https://open.spotify.com/playlist/37i9dQZF1CHWPLEJxCaAr7 |
| Teen & Young Adult | https://open.spotify.com/playlist/37i9dQZF1CHSLN7nOfbh3E |
| Business & Careers | https://open.spotify.com/playlist/37i9dQZF1CHND3tYQBenWx |
| Entertainment Biography | https://open.spotify.com/playlist/37i9dQZF1CHCmXPRIw1m4n |
| SciFi & Fantasy | https://open.spotify.com/playlist/37i9dQZF1CHShViq8tHUSO |
| Self-Help | https://open.spotify.com/playlist/37i9dQZF1CHSGVWsnPOKrL |

## How it works

1. **Firecrawl** scrapes each playlist page (`open.spotify.com/playlist/<id>`)
2. **Parser** extracts N, Title, Author from the rendered markdown
3. **Gemini** (with Google Search grounding) looks up first-edition print publisher and date per book
4. **Imprint → Big-5 mapping** classifies each book as Penguin Random House / HarperCollins / Simon & Schuster / Macmillan / Hachette / Bloomsbury / Independent
5. **Backlist flag** marks books >12 months old vs today
6. **PDF + xlsx** rendered per playlist with rank-change column (NEW / ↑n / ↓n / —)

## Local usage

```bash
python3 refresh_all.py                    # process all playlists
python3 refresh_one.py <id> --tab "Name"  # single playlist
python3 refresh_one.py <id> --cache       # skip Firecrawl, use cached markdown
```

Credentials read from `.creds/gemini_api_key.txt` and `.creds/google_books_key.txt`. Firecrawl key comes from the `firecrawl-key` CLI swapper or `FIRECRAWL_API_KEY` env var.

## CI (GitHub Actions)

The `.github/workflows/weekly_refresh.yml` workflow runs every Monday 13:00 UTC. It:
- Loads `GEMINI_API_KEY` and `FIRECRAWL_API_KEY` from repo secrets
- Runs `refresh_all.py`
- Commits the updated `audiobooks.xlsx` and PDFs back to the repo
- Uploads PDFs as a workflow artifact (downloadable from the Actions tab for 90 days)

Trigger a manual refresh anytime from the **Actions** tab → **Weekly audiobook refresh** → **Run workflow**.

## Output

- `audiobooks.xlsx` — every tab refreshed in place; manually-curated tabs (Top, Romance) are also refreshed but were previously preserved
- `output/pdfs/audiobooks_<Tab>_YYYY-MM-DD.pdf` — one PDF per chart per refresh, committed to the repo
