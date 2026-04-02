# Sprava faktur (GUI)

Aplikace obsahuje:
- moderni graficke rozhrani
- spravu fakturujicich firem (vice firem, vychozi firma)
- spravu odberatelu vcetne editace
- automaticke cislovani faktur ve formatu `DDMM_XXXX`
- export faktury do HTML i PDF
- QR kod na platbu (SPD)
- import faktur z Excelu (a automaticke doplneni odberatelu)
- rocni prehledy (vsechny roky + detail mesicu + aktualni rok)

## Instalace Pythonu

```powershell
winget install -e --id Python.Python.3.12
```

## Instalace knihoven

```powershell
python -m pip install reportlab qrcode[pil] openpyxl
```

## Spusteni

```powershell
cd "C:\Users\aabon.sedlmi\Documents\AI chat"
python invoice_manager_gui.py
```

## Import z Excelu

V zalozce `Faktury` klikni na `Import Excel` a vyber `.xlsx` nebo `.xlsm`.

### Minimalni sloupec
- `customer_name`

### Doporucene sloupce
- `customer_name`
- `customer_email`
- `customer_address`
- `issue_date` (`YYYY-MM-DD` nebo `DD.MM.YYYY`)
- `due_date`
- `status` (`draft|sent|paid|overdue`)
- `currency` (napr. `CZK`)
- `note`
- `invoice_number` (volitelne, kdyz neni, vygeneruje se `DDMM_XXXX`)
- `item_description`
- `quantity`
- `unit_price`
- `vat_rate`

Kazdy radek reprezentuje jednu polozku faktury. Radky se stejnym `invoice_number` se seskupuji do jedne faktury.

## Dulezite soubory

- GUI: `invoice_manager_gui.py`
- Backend: `invoice_backend.py`
- Databaze: `invoices.db`

### Alternativni format Excelu (tvuj)
- Jeden soubor Excel.
- Kazdy list/stranka = jedna faktura.
- Aplikace se pokusi nacist pole jako: odberatel, adresa, email, datum vystaveni, datum splatnosti, mena, cislo faktury, stav a tabulku polozek (popis/mnozstvi/cena/DPH).
- Pokud cislo faktury chybi, vygeneruje se automaticky DDMM_XXXX.


## Dalsi nove funkce
- Smazani faktury z GUI (vcetne polozek).
- Export vsech dat do JSON (Export dat).
- Import vsech dat z JSON (Import dat) - prepise aktualni data.
- Opraveny import z listu: label Odběratel: je nyni nacitan i kdyz je hodnota v sousedni bunce.


- Odberatel nyní obsahuje: nazev, adresa, telefon, email, ICO, DIC.
- V zalozce Odberatele je hromadne mazani vybranych zaznamu (bez navazanych faktur).

