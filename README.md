# Family finance tracker

Double-click **run.bat** and pick an option. (First time: install Python from python.org and tick "Add python.exe to PATH". run.bat installs the rest.)

## Monthly routine
1. Download the new exports (Leumi, Mizrahi, Isracard per card, Max) into this folder or into `exports/`.
   Overlapping date ranges are fine - rows already loaded are skipped.
2. **1 - Import.** `output/review.xlsx` opens with everything that needs your answer.
3. Easiest first: open `output/shops.xlsx` - one row per shop still waiting. Pick a category in the last
   column ("קטגוריה"); it applies to every transaction of that shop.
   Editing in Google Sheets? Download it as CSV into this folder with a name starting with `shops` - Apply reads it. Then use review.xlsx for what's left.
   In review.xlsx: pick a category in the last column. Yellow cells are suggestions from your past answers.
   Answer one row per shop: after Apply, the other rows of the same shop are categorized automatically.
   Bit, PayBox, cash, checks and transfers are always asked, one by one.
   To split a payment: copy the row (same ID) and change the amounts so they add up to the original.
4. Save and close the file(s), then **2 - Apply**. Repeat 3-4 until nothing is left.
5. **3 - Report** (all years) or **4 - Report** for one year -> `output/report.xlsx`.

## Files you can edit (in Excel)
- `config/categories.csv` - your categories and sub-categories (type: הוצאה / הכנסה).
- `config/rules.csv` - fixed rules, checked top to bottom, first match wins.
  - התאמה: מכיל (contains), מילה (whole word), מדויק (exact description).
  - חשבון או כרטיס: optional, e.g. 1234 to apply a rule only to that card.
  - קטגוריה "התעלם" = never counted; "שאל תמיד" = always ask me.
- `config/card_hints.csv` - Max's own category -> the category suggested to you (yellow). Only a suggestion.
- `config/accounts.csv` - add a nickname per account/card (e.g. בן זוג 1, בן זוג 2).

## How it decides
- Monthly card debits in the bank (מקס, ישראכרט, כאל) are skipped; the card files hold the detail.
  The "בדיקות" sheet in the report checks that card files add up to the bank debits and lists cards with no file.
- Equal amounts going out of one of your accounts and into the other within 3 days are skipped as own transfers.
- Rules first. Then your past answers: if at least 95% of your answers for that shop agree, it is applied.
  Otherwise you are asked, with the most common past answer as a suggestion.
- Installments are counted in the month they are charged; other card purchases in the purchase month.

## Data
Everything is CSV in `data/` (transactions, splits, memory, import log). Every run backs it up to `backups/`.
You can fix a category directly in `data/transactions.csv` (set status to `user`).

## Yearly file ("רישום כלכלי YYYY.xlsx")
Report (option 3) also writes `output/רישום כלכלי <year>.xlsx` in the layout of your manual file:
"כללי" (expenses, income, difference, maaser, money on the 15th), a sheet per category with
sub-categories, item lists for שונות and מתנות ותרומות, and all transactions.
The file is rebuilt on every run, so anything you want to keep goes in config:
- `config/balances.csv` - money on the 15th (שם, חודש like 2026-09, סכום).
- `config/notes.csv` - notes per month (חודש, אזור = הוצאות/הכנסות, הערה).
- `config/recurring.csv` - fixed monthly amounts that don't show up in the bank (rent, nanny...).
A salary paid on the 25th or later is counted in the next month.
