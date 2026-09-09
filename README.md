# Duplex Name Tag Mail Merge

Print double-sided, laminated name tags — 6 per sheet, front and back automatically
lined up — using nothing but Word mail merge and an Excel formula. No macros,
no add-ins.

## The problem this solves

A common way to make cheap, durable name tags is to mail-merge a 2×3 grid of
tags onto cardstock, print it double-sided, laminate the whole sheet, then cut
it into six individual tags. That way each tag is readable from both sides.

The catch: for the back of each tag to show the *same* person's name as the
front, the back page can't just repeat the same left-to-right order as the
front. Depending on which edge your printer flips on, the physical position of
each cell moves when the sheet turns over — so the back page has to be printed
in a re-shuffled order for everything to land in the right place.

This repo contains a Word template with the merge fields already laid out, and
an Excel workbook that reorders your name list into the correct print order
automatically — so you can maintain a plain, ordinary list of names and never
have to think about the shuffling by hand.

## How the pattern works

Assume your printer duplexes by flipping on the **long edge** (the standard
default for portrait pages on most printers — verify this before a large print
run; see [Adapting this](#adapting-this-to-your-printer) if yours flips on the
short edge instead).

For a 2-column × 3-row sheet, a long-edge flip mirrors each row left-to-right.
So if the front of the sheet is filled in reading order:

| Front (reading order) | Position 1 | Position 2 |
|---|---|---|
| Row 1 | Person A | Person B |
| Row 2 | Person C | Person D |
| Row 3 | Person E | Person F |

...the back of that same sheet needs to be *merged* in this order so each
name lands behind its own front:

| Back (merge order) | Position 1 | Position 2 |
|---|---|---|
| Row 1 | Person B | Person A |
| Row 2 | Person D | Person C |
| Row 3 | Person F | Person E |

In other words: every pair of positions on a row swaps. For six people, the
12 rows fed into the mail merge (6 front + 6 back) come out as:

```
A, B, C, D, E, F,   B, A, D, C, F, E
```

...and this repeats for every group of 6 people in your list. The Excel
workbook generates this automatically from a plain list — you maintain the
plain list, it computes the merge order.

## What's in this repo

```
templates/
  Name_Tag_Template.docx        Word mail-merge template, 6 tags per sheet
  Name_Tag_Roster_Workbook.xlsx Roster + auto-generated Mail Merge List
```

The Word template has a `[ YOUR LOGO / ORGANIZATION NAME ]` placeholder in
each cell — swap that for your own logo/wordmark once, and it carries across
every tag.

The workbook has two tabs:

- **Roster** — a plain list of names, one row per person. This is the only
  tab you edit.
- **Mail Merge List** — entirely formulas. It reads Roster and produces the
  front/back merge order shown above, 12 rows per 6 people, along with an
  `Include` column (Yes/No) to filter out unused placeholder rows, and a
  `Notes` column that says which physical sheet/side/position each row maps
  to, for troubleshooting.

Full step-by-step usage instructions are on the **Instructions** tab of the
workbook itself.

## Quick start

1. Open `Name_Tag_Roster_Workbook.xlsx` and enter your names on the **Roster**
   tab (First, Last — one row per person).
2. Open `Name_Tag_Template.docx` in Word. Replace the logo placeholder text
   with your own logo/name. Go to **Mailings → Select Recipients → Use an
   Existing List** and point it at the workbook, sheet `Mail Merge List`.
3. **Mailings → Edit Recipient List** → filter `Include` = `Yes` (hides
   unused placeholder rows built in for future growth; a one-time setting).
4. **Mailings → Finish & Merge → Edit Individual Documents → All.** Proofread
   the merged document on screen before printing.
5. Print two-sided, flip on the long edge. **Print one test sheet on plain
   paper first** and hold it up to the light to confirm names line up
   front-to-back before running a full batch on cardstock.
6. Laminate, trim along the grid lines, and attach lanyards/clips.

## Adapting this

- **Different grid size (not 6-per-sheet):** the swap pattern generalizes —
  every row of the grid swaps left-to-right for a long-edge flip, regardless
  of how many rows there are. Adjust the `CHOOSE(...)` list in the workbook's
  formulas to match your row/column count.
- **Printer flips on the short edge instead:** a short-edge flip mirrors
  top-to-bottom instead of left-to-right — rows swap with each other (row 1 ↔
  row 3, with row 2 unchanged for a 3-row grid) rather than columns swapping
  within a row. Replace the pattern constant in the formulas from
  `1,2,3,4,5,6,2,1,4,3,6,5` to `1,2,3,4,5,6,5,6,3,4,1,2` to get that mapping.
  **Test on plain paper either way** — it's the cheapest way to confirm which
  one your printer actually needs.
- **More than 240 people:** select the last row of `Mail Merge List` and
  copy its formulas down as far as you need. They reference their own row
  number, so they don't need to be edited, just extended.

## A maintenance tip that will save you a reprint

Because grouping is based on a person's *position* in the Roster list (not a
stable ID), reordering, inserting, or deleting rows in the middle of an
existing roster shifts everyone below into different groups of 6 — meaning
any tags already printed and laminated stop matching. Always add new people
to the **bottom** of the list; if someone leaves, clear their name but leave
the row in place rather than deleting it.

## License

MIT — see [LICENSE](LICENSE). Use, adapt, and share freely.
