# Spotify Audiobook Chart Tracker

Weekly snapshot of Spotify's editorial audiobook playlists into a tracked spreadsheet + per-chart PDFs, with rank-change deltas vs the previous refresh.

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
