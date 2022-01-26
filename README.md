# Overwatch Data Parser

## Installation :: Requires Python 3.10
1. Download all files
2. Install packages in `requirements.txt`. I reccomend using a virtual environment.
3. Get your Google Sheets API service account credentials as a JSON file. Guide: https://docs.gspread.org/en/latest/oauth2.html
4. Rename the credentials.json file to `creds.json` and place it in the root directory of this project.
5. Make a copy of this google sheet for formatted data:  https://docs.google.com/spreadsheets/d/1255FPypvcxPGqN7mP-NMMkwN15PTlx6RxDiQS6A_NCQ/edit?usp=sharing
6. Invite the service account to the Google Sheet you would like to dump to as an EDITOR (See guide above). 
7. Open `main.py` in your text editor. Scroll down to `manager = SheetManager("creds.json", "Sens: Season 32")` near the bottom of the file, and change `"Sens: Season 32"` to the name of your google sheet.
8. run `main.py`

## Example
https://docs.google.com/spreadsheets/d/1255FPypvcxPGqN7mP-NMMkwN15PTlx6RxDiQS6A_NCQ/edit?usp=sharing

## Licence
[MIT](https://choosealicense.com/licenses/mit/)