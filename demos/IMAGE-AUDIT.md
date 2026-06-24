# Demo Image Asset Audit

> Compliance + metadata inventory of images under `demos/`. Generated 2026-06-24.
> Dimensions/size/format are mechanical (via `image-size`); **Risk** and **Source** are human assessment.
>
> ⚠️ These assets are **demo-only**. They are NOT part of the published package
> (`files: ["dist","types"]` excludes them) — see the note at the bottom. This audit
> covers the risk of **hosting the demo app** or **sharing the repo source**.
>
> Caveat: removing a file from the working tree does not remove it from **git history**.
> A full scrub requires history rewrite (git-filter-repo / BFG).

## 🔴 High risk — identifiable third-party IP / trademarks

| File | Dimensions | Size | Fmt | Why unsafe |
| --- | --- | ---: | --- | --- |
| `common/images/starlabs_logo.png` | 600×270 | 36.1 KB | png | **S.T.A.R. Labs** — DC Comics fictional brand (trademark/IP) |
| `common/images/starlabs_bkgd.jpg` | 1920×1080 | 300.6 KB | jpg | Matching S.T.A.R. Labs background (DC Comics) |
| `common/images/chicago_bean_bohne.jpg` | 256×192 | 20.4 KB | jpg | **Cloud Gate** sculpture (Anish Kapoor) — commercial photo copyright is actively enforced |
| `common/images/nyc-subway.png` | 900×600 | 301.5 KB | png | MTA NYC subway map — copyrighted transit map |
| `common/images/tokyo-subway-route-map.jpg` | 400×279 | 69.0 KB | jpg | Tokyo Metro map — copyrighted transit map |
| `common/images/krita_square.jpg` | 500×500 | 47.4 KB | jpg | Krita mascot **"Kiki"** (Tyson Tan) — character art, specific license |
| `common/images/krita_splashscreen.jpeg` | 2276×1280 | 448.3 KB | jpeg | Krita splash artwork — artist copyright |
| `common/images/wiki-example.jpg` | 275×297 | 27.0 KB | jpg | Wikipedia placeholder — Wikimedia globe trademark |
| `browser/images/github.svg` | 16×16 | 0.7 KB | svg | GitHub Octocat — trademark, brand guidelines |
| `browser/images/mastodon.svg` | 16×16 | 0.9 KB | svg | Mastodon logo — trademark |
| `common/images/mastodon-logo-purple.svg` | 75×79 | 2.4 KB | svg | Mastodon logo — trademark |
| `common/images/pixelfed_icon.svg` | 75×75 | 14.9 KB | svg | Pixelfed logo — trademark |
| `common/images/fediverse_actpub.png` | 1114×1200 | 975.1 KB | png | ActivityPub / Fediverse branding |
| `common/images/fediverse_tree.jpg` | 1000×1001 | 167.6 KB | jpg | Fediverse branding |
| `common/images/cc_logo.jpg` | 960×720 | 42.1 KB | jpg | Creative Commons brand asset (trademark) |
| `common/images/cc_symbols_trans.png` | 800×302 | 30.2 KB | png | Creative Commons symbols (trademark) |
| `common/images/cc_license_comp.png` | 800×473 | 239.2 KB | png | Creative Commons license chart (trademark) |
| `common/images/cc_copyremix.gif` | 351×359 | 41.2 KB | gif | Creative Commons campaign asset |
| `common/images/cc_dj.gif` | 208×166 | 5.4 KB | gif | Creative Commons campaign asset |

## 🟡 Medium risk — unattributed photos / GIFs (no license on file)

| File | Dimensions | Size | Fmt | Notes |
| --- | --- | ---: | --- | --- |
| `common/images/image2.jpg` | 960×720 | 102.6 KB | jpg | Stock-style landscape photo, unattributed |
| `common/images/sydney_harbour_bridge_night.jpg` | 1920×479 | 152.9 KB | jpg | Night photo, unattributed |
| `common/images/trippy.gif` | 500×500 | **2004.2 KB** | gif | Unknown web GIF — also 2 MB of bloat |
| `common/images/anim_campfire.gif` | 128×128 | 3.9 KB | gif | Unknown web GIF |
| `common/images/title_bkgd.jpg` | 4800×2700 | **1853.1 KB** | jpg | Unattributed background — also 1.8 MB of bloat |
| `common/images/title_bkgd_alt.jpg` | 1200×600 | 68.3 KB | jpg | Unattributed background |
| `common/images/cover_audio.png` | 937×937 | 735.5 KB | png | Media cover — possible copyrighted still |
| `common/images/cover_video_16x9.png` | 1280×720 | 681.7 KB | png | Media cover — possible copyrighted still |
| `common/images/sample-hd-m4v-cover.png` | 920×505 | 734.8 KB | png | Media cover — possible copyrighted still |
| `common/images/video-mp4-thumb.png` | 518×387 | 272.1 KB | png | Media thumbnail — possible copyrighted still |
| `common/images/peace4.png` | 400×400 | 188.1 KB | png | Unattributed graphic |
| `common/images/unite.png` | 331×291 | 170.1 KB | png | Unattributed graphic |

## 🟢 Likely safe — project's own / generic test assets

| File | Dimensions | Size | Fmt |
| --- | --- | ---: | --- |
| `browser/images/favicon.png` | 480×349 | 15.4 KB | png |
| `browser/images/favicon-16x16.png` | 16×16 | 0.5 KB | png |
| `browser/images/favicon-32x32.png` | 32×32 | 1.3 KB | png |
| `browser/images/html2pptx.png` | 1712×836 | 672.8 KB | png |
| `browser/images/slide-master.png` | 2088×1248 | 859.9 KB | png |
| `browser/images/info-circle.svg` | 16×16 | 0.4 KB | svg |
| `common/images/lock-green.svg` | 500×500 | 15.4 KB | svg |
| `common/images/logo_square.png` | 210×210 | 27.9 KB | png |
| `common/images/logo_square_25.png` | 25×25 | 1.4 KB | png |
| `common/images/logo_square_50.png` | 50×50 | 4.5 KB | png |
| `common/images/sample_logo.png` | 598×302 | 34.0 KB | png |
| `common/images/png-gradient-hex.png` | 812×621 | 87.5 KB | png |
| `common/images/play-button.png` | 1920×1383 | 54.5 KB | png |
| `common/images/title_bkgd.png` | 1400×900 | 7.5 KB | png |
| `common/images/UPPERCASE.PNG` | 301×167 | 9.2 KB | png |
| `common/images/brokenImage.gif` | 100×119 | 1.3 KB | gif |
| `common/images/brokenImage.png` | 200×238 | 3.8 KB | png |
| `node/assets/image.png` | 192×192 | 8.4 KB | png |

## 🟢 Vite/React scaffold defaults (third-party marks, standard)

| File | Dimensions | Size | Fmt |
| --- | --- | ---: | --- |
| `vite-demo/public/vite.svg` | 32×32 | 1.5 KB | svg |
| `vite-demo/src/assets/react.svg` | 36×32 | 4.0 KB | svg |
| `vite-demo/src/assets/logo.png` | 30×30 | 0.6 KB | png |

## ⚠️ Media (audio/video) — higher copyright risk than images

`demos/common/media/` holds `earth-big.mp4` (17 MB), `sample-hd.m4v`, `sample.{avi,mov,mpg,m4v,mp4,mp3,wav,aif}`.
Audio/video most likely carries copyrighted **music/footage** — not dimension-audited here, but should get the same provenance review.

## Distribution note

The published package contains only `dist/`, `types/`, `README.md`, `LICENSE` (verified via `npm pack`).
**None of these assets ship to package consumers.** Risk applies only to hosting the demo app or
sharing the repo source. README's banner is a remote hotlink to upstream, not a bundled file.
