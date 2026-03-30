# Invoice Converter — Excel → UBL XML

Web app that converts Business Central posted sales invoice exports (`.xlsx`) into UBL 2.1 XML for the Serbian eFaktura system (EN 16931 / mfin.gov.rs 2022 schema).

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit web UI |
| `excel_to_ubl_xml.py` | Conversion logic |
| `requirements.txt` | Python dependencies |

---

## ⚙️ Before deploying — set your seller details

Open `excel_to_ubl_xml.py` and fill in the `SELLER` dictionary near the top:

```python
SELLER = {
    "pib":           "123456789",          # your 9-digit PIB
    "name":          "Vaša Firma d.o.o.",
    "street":        "Ulica i broj",
    "city":          "Beograd",
    "post_code":     "11000",
    "country":       "RS",
    "mb":            "12345678",           # your 8-digit matični broj
    "email":         "fakture@vasa-firma.rs",
    "bank_account":  "160-123456789-12",   # your tekući račun
}
```

---

## 🚀 Deploy to Streamlit Community Cloud (free, ~5 minutes)

### Step 1 — Push to GitHub

1. Create a free account at [github.com](https://github.com) if you don't have one
2. Create a **new repository** (e.g. `invoice-converter`) — set it to **Private** if you want
3. Upload all three files into the repository:
   - `app.py`
   - `excel_to_ubl_xml.py`
   - `requirements.txt`

### Step 2 — Deploy on Streamlit

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with your GitHub account
2. Click **"New app"**
3. Select your repository and branch (`main`)
4. Set **Main file path** to `app.py`
5. Click **Deploy**

Streamlit will install dependencies and start the app. In ~2 minutes you get a public URL like:
```
https://your-app-name.streamlit.app
```

Share that URL with anyone in the company — they open it in any browser, upload their `.xlsx`, and download the `.xml`. No installation required.

---

## Updating the app

To update the conversion logic or seller details, just edit the files in GitHub (or push a new commit). Streamlit Cloud redeploys automatically within seconds.
