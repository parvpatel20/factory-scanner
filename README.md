# Factory Scanner

Upload up to 5 factory notebook images, extract the row values with Groq vision, review and correct each table, then download separate Excel files with row totals, column totals, and a grand total.

## Setup

Install Python 3.8 or newer, then install the dependencies:

```bash
pip install -r requirements.txt
```

The app includes a default Groq key in the local backend. No key entry is shown in the browser.

## Run

### Mac / Linux

```bash
python3 server.py
```

You can also run:

```bash
./start.sh
```

### Windows

```bat
python server.py
```

You can also double-click `start.bat`.

Open the app at:

```text
http://localhost:5000
```

## Use

1. Upload 1 to 5 clear photos of the notebook or factory sheets.
2. Click `Extract selected images`.
3. Open any image page from the image list.
4. Review the editable table and correct any wrong cells.
5. The row totals, column totals, and grand total update as you type.
6. Click `Download this Excel` for one image, or `Download all Excels` to get a zip containing one workbook per extracted image.

## Configuration

`GROQ_API_KEY` is optional. If set, it overrides the default key configured in the server.

`GROQ_MODEL` is optional. By default the app uses:

```text
meta-llama/llama-4-scout-17b-16e-instruct
```

## Tips

- Better lighting and a straight-on photo improve extraction accuracy.
- The app compresses large images locally before sending them to Groq.
- Use `Create blank table` when you want to enter values manually.
- Use the back buttons to return to upload, the image list, or a previous image page without losing completed tables.
- Downloaded Excel files contain formulas for totals, so later edits in Excel keep totals correct.
