# Logger Data Separator

Turns bulk Elitech / MU-01 **temperature + humidity** logger reports (PDF or Excel) into one workbook with two sheets:

- **Temperature** — `DATE | Time | DL-1 | DL-2 | …`
- **Humidity** — same timestamps and logger columns

Each file named like `DL 01` becomes column `DL-1`. Readings are aligned by date and time.

A full folder of ~270 loggers takes about 3 to 4 minutes and peaks around 80 MB of memory.

## Run on your PC

Double-click `start.bat`, or:

```bat
python -m pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5000, paste the folder that holds the DL files, click **Check folder**, then **Separate all DLs**.

To process a folder without the browser:

```bat
python process_folder.py "C:\path\to\DL folder"
```

## Run on Render

The repo already contains `render.yaml`, so Render configures itself.

1. Push this folder to a GitHub repository.
2. In Render: **New → Blueprint**, pick the repository, and confirm.
3. Optionally set `APP_PASSWORD` to require a shared password. Leave it unset and the site is open to anyone with the link.

Render runs with `CLOUD_MODE=1`, which hides the folder-path box, because the server cannot see your PC's folders. Files are uploaded from the browser 15 at a time so a large batch does not fail as one huge request, and uploaded files are deleted as soon as the job finishes.

Keep the browser tab open while a job runs; progress is polled from the server.

### Environment variables

| Variable | Purpose |
|---|---|
| `CLOUD_MODE` | `1` hides the folder-path input and binds to all interfaces |
| `APP_PASSWORD` | Shared password. Unset means no login |
| `SECRET_KEY` | Signs the login session; Render generates one |
| `DATA_DIR` | Where `output/` and `uploads/` live |
| `PORT` | Port to listen on |

## Files

| File | Purpose |
|---|---|
| `app.py` | Web app, upload handling, job runner |
| `parser.py` | PDF and Excel reading, timeline merge |
| `excel_export.py` | Writes the two-sheet workbook |
| `process_folder.py` | Command-line run over a local folder |
| `check_fill.py` | Reports how many readings landed in each column |
| `test_cloud.ps1` | End-to-end check of the cloud upload flow |
