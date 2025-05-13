# Create Volusion Loadsheet
Creates GUI translating product information from e-commerce site Volusion.

## How To Create Executable File

### Step 1: Download Product CSV File from Volusion and Python File From Repository (create_volusion_loadsheet.py)

If using sample csv file, download "volusion_sample_data.csv"

### Step 2: Download the following onto your computer:

#### Python (Programming Language): 

How To Download:

https://www.python.org/downloads/

Run this command in terminal to see if downloaded properly:

```bash
python --version
```

pyinstaller --onefile --noconsole create_volusion_loadsheet.py

#### Pip (Python Package Manager):

How To Download:

pip install pyinstaller

Run this command in terminal to see if downloaded properly:

```bash
pip --version
```

#### Pyinstaller (Converts Python Scripts Into Executable Files):

How To Download:

pip download pyinstaller

Run this command in terminal to see if downloaded properly:

```bash
pyinstaller --version
```
### Step 3:  In terminal, go to the same folder/directory as the Python file and enter the following command:

```bash
pyinstaller --onefile --noconsole create_volusion_loadsheet.py
```
