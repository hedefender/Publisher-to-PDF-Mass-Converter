# MS Publisher to PDF Mass Converter

A lightweight, standalone tool with a graphical user interface (GUI) for mass converting Microsoft Publisher (`.pub`) files to PDF format. 

Created to assist IT professionals and system administrators ahead of the MS Publisher end of support, automating the tedious process of exporting hundreds of files.

## About MS Publisher End of Support
Microsoft has officially announced that support for Microsoft Publisher will end on **October 13, 2026**. After this date, the application will not be included in new Microsoft 365 subscriptions and will not receive security updates. This tool was created to facilitate the transition and preservation of files.

* **Official Announcement:** [Microsoft Publisher will no longer be supported after October 2026](https://support.microsoft.com/en-us/office/microsoft-publisher-will-no-longer-be-supported-after-october-2026-87807759-45cc-4c28-98f9-58a36c84f501)

## Key Features
* **Automatic Scanning:** Searches for `.pub` files in the selected root folder and all its subfolders.
* **Structure Preservation:** The generated `.pdf` file is saved right next to the original `.pub` file, keeping your file server/disk folder structure intact.
* **Safe Conversion:** Uses the MS Publisher COM objects locally in the background, ensuring 100% formatting fidelity without altering the original files.
* **Temporary Workspace:** File processing occurs in the Windows temporary folder (temp) to prevent file locking errors.

## Prerequisites
For the executable (or the source code) to work properly, **Microsoft Publisher must be installed** on the system running the application.

## Usage Instructions (For End Users)
1. Download the `Converter.exe` executable from the Releases page.
2. Run the application (no installation required).
3. Click the **Select Folder & Start** button and choose the root folder containing your files.
4. The application will scan the folders, start the background conversion, and notify you upon completion.

## Development Instructions (For Developers)
If you want to run the code via Python or modify it:

1. Clone the repository:
   ```bash
   git clone [https://github.com/yourusername/publisher-to-pdf-converter.git](https://github.com/yourusername/publisher-to-pdf-converter.git)

2. Create and activate a virtual environment:
```bash
python -m venv .venv
.\.venv\Scripts\activate

```


3. Install the required libraries:
```bash
pip install pywin32 pyinstaller

```


4. Run the application:
```bash
python Converter.py

```


5. To build your own standalone `.exe`:
```bash
pyinstaller --noconsole --onefile Converter.py

```



License
This tool is available for free as open-source software. It is provided "as-is", without any warranty for its operation.

Creator: Stefanos Paraskevas



# MS Publisher to PDF Mass Converter

Ένα ελαφρύ, αυτόνομο εργαλείο με γραφικό περιβάλλον (GUI) για τη μαζική μετατροπή αρχείων Microsoft Publisher (`.pub`) σε μορφή PDF. 

Δημιουργήθηκε για να διευκολύνει επαγγελματίες πληροφορικής και διαχειριστές συστημάτων ενόψει της λήξης υποστήριξης του MS Publisher, αυτοματοποιώντας την "επίπονη" διαδικασία εξαγωγής εκατοντάδων αρχείων.

## Σχετικά με την απόσυρση του MS Publisher
Η Microsoft έχει ανακοινώσει επίσημα ότι η υποστήριξη για το Microsoft Publisher θα τερματιστεί στις **13 Οκτωβρίου 2026**. Μετά από αυτή την ημερομηνία, η εφαρμογή δεν θα περιλαμβάνεται στις νέες συνδρομές Microsoft 365 και δεν θα λαμβάνει ενημερώσεις ασφαλείας. Το παρόν εργαλείο δημιουργήθηκε για να διευκολύνει τη μετάβαση και τη διατήρηση των αρχείων.

* **Επίσημη Ανακοίνωση:** [Microsoft Publisher will no longer be supported after October 2026](https://support.microsoft.com/en-us/office/microsoft-publisher-will-no-longer-be-supported-after-october-2026-87807759-45cc-4c28-98f9-58a36c84f501)

## Βασικά Χαρακτηριστικά
* **Αυτόματη Σάρωση:** Αναζητά αρχεία `.pub` στον φάκελο που θα επιλέξετε, καθώς και σε όλους τους υποφακέλους του.
* **Διατήρηση Δομής:** Το παραγόμενο αρχείο `.pdf` αποθηκεύεται ακριβώς δίπλα στο αρχικό αρχείο `.pub`, διατηρώντας ανέπαφη τη δομή των φακέλων του file server/δίσκου σας.
* **Ασφαλής Μετατροπή:** Χρησιμοποιεί τοπικά τα COM objects του ίδιου του MS Publisher στο παρασκήνιο, εξασφαλίζοντας 100% πιστότητα στη μορφοποίηση του εγγράφου χωρίς να επεμβαίνει στα αρχικά αρχεία.
* **Προσωρινός Χώρος Εργασίας:** Η επεξεργασία του κάθε αρχείου γίνεται στον προσωρινό φάκελο (temp) των Windows για την αποφυγή σφαλμάτων κλειδώματος αρχείων.

## Προαπαιτούμενα
Για να λειτουργήσει σωστά το εκτελέσιμο (ή ο κώδικας), **πρέπει υποχρεωτικά να είναι εγκατεστημένο το Microsoft Publisher** στο σύστημα που εκτελείται η εφαρμογή. 

## Οδηγίες Χρήσης (Για τελικούς χρήστες)
1. Κατεβάστε το εκτελέσιμο αρχείο `Converter.exe` από τα Releases.
2. Τρέξτε την εφαρμογή (δεν απαιτείται εγκατάσταση).
3. Πατήστε το κουμπί **Select Folder & Start** και επιλέξτε τον κεντρικό φάκελο που περιέχει τα αρχεία σας.
4. Η εφαρμογή θα σαρώσει τους φακέλους, θα ξεκινήσει τη μετατροπή στο παρασκήνιο και θα σας ενημερώσει μόλις ολοκληρωθεί.

## Οδηγίες Ανάπτυξης (Για developers)
Αν θέλετε να τρέξετε τον κώδικα μέσω Python ή να τον τροποποιήσετε:

1. Κάντε clone το repository:
   ```bash
   git clone [https://github.com/yourusername/publisher-to-pdf-converter.git](https://github.com/yourusername/publisher-to-pdf-converter.git)

2. Δημιουργήστε και ενεργοποιήστε ένα εικονικό περιβάλλον:
```bash
python -m venv .venv
.\.venv\Scripts\activate

```


3. Εγκαταστήστε τις απαιτούμενες βιβλιοθήκες:
```bash
pip install pywin32 pyinstaller

```


4. Εκτελέστε την εφαρμογή:
```bash
python Converter.py

```


5. Για εξαγωγή σε δικό σας αυτόνομο `.exe`:
```bash
pyinstaller --noconsole --onefile Converter.py

```



## Άδεια Χρήσης (License)

Το παρόν εργαλείο διατίθεται δωρεάν ως open-source λογισμικό. Παρέχεται "ως έχει" (as-is), χωρίς καμία εγγύηση για τη λειτουργία του.

**Δημιουργός:** Στέφανος Παρασκευάς
