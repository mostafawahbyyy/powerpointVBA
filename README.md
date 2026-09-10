# PowerPoint VBA Macros

Small VBA macros for automating repetitive PowerPoint formatting tasks.

## StandardizeTitles

Loops through every slide in the active presentation and applies consistent
formatting to the title placeholder: font, size, weight, colour, alignment,
and position.

Slides without a title placeholder are skipped rather than causing a runtime
error. A summary message reports how many titles were formatted and how many
slides were skipped.

### Usage

1. Open the presentation
2. Developer tab → Visual Basic → Insert → Module
3. Paste the contents of `StandardizeTitles.bas`
4. Run `StandardizeTitles`

### Note

Written against the PowerPoint object model but not executed — I don't
currently have an Office license to test against.
