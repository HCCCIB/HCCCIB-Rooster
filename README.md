# HCCC Dienstenrooster 2026

Live rooster app voor HCCC Installatiebeheer – Schiphol.

🔗 **[Open rooster app](https://hcccib.github.io/HCCCIB-Rooster/)**

---

## Rooster bijwerken

1. Ga naar de **[`content/`](./content/)** map in deze repository
2. Klik op het bestaande Excel bestand → **"..."** → **"Replace file"**  
   *(of upload een nieuw `.xlsx` bestand)*
3. Commit de upload — de app wordt **automatisch bijgewerkt** binnen ~1 minuut

> De GitHub Action leest de Excel, extraheert de dienstdata en bouwt een nieuwe `index.html`.  
> Je hoeft verder niks te doen.

---

## Bestandsstructuur

```
/
├── index.html              ← Live app (automatisch gegenereerd, niet handmatig aanpassen)
├── index.template.html     ← Basis HTML zonder roosterdata (bron voor de builder)
├── content/
│   └── Dienstenrooster-HCCC-2026.xlsx   ← Upload hier je nieuwe Excel
└── scripts/
    └── build.py            ← Python script dat Excel → index.html bouwt
```

---

## Dienstcodes

| Code | Dienst | Tijd |
|------|--------|------|
| o | Ochtend | 05:45–14:15 |
| m | Middag | 13:45–22:15 |
| n | Nacht | 21:45–06:15 |
| mx | Mid-extended | 10:30–19:00 |
| x | Dagdienst | 08:00–16:30 |
| v | Verlof | – |
| z | Ziek | – |
| f | Feestdag | – |
