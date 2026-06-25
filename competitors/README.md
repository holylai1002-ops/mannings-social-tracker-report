# Fanpage Karma Screenshots

## Monthly Workflow

1. Login to **https://www.fanpagekarma.com**
   - Account: Sindy Google (focusae2012@gmail.com)

2. Navigate to **Mannings > Competitor Analysis**

3. For each chart/table you want to capture:
   - Take a screenshot (`Win + Shift + S` or `PrtScn`)
   - Save as PNG to: `screenshots/{year}_{month}/`

## Naming Convention

Files should match names in `screenshots_config.yaml`:

```
screenshots/
  2026_05/
    competitor_overview.png        <- Competitor Overview
    competitor_facebook.png        <- Facebook Competitor Analysis
    competitor_instagram.png       <- Instagram Competitor Analysis
    competitor_engagement.png      <- Engagement Comparison
    competitor_growth.png          <- Follower Growth
  2026_06/
    ...
```

## Adding Custom Screenshots

Any additional PNG files dropped in a period folder will also be displayed
(using the filename as the label). No config change needed.

## Example

```powershell
# After taking screenshots, just drop them in the folder:
Copy-Item "$env:USERPROFILE\Desktop\screenshot1.png" `
          ".\screenshots\2026_05\competitor_overview.png"
```
