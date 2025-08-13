# Create Volusion Loadsheet

This repository creates a GUI translating product information from e-commerce site Volusion. The resulting load sheet contains the following information:

* **Part #**: taken from "productcode" field
* **Full Title**: taken from "productname" field
* **Parent #**: parent product id taken from "ischildofproductcode" field
* **Parent Title**: title of parent product based on "Parent #"
* **Retail Price**: taken from "productprice" field
* **Length (in)**: taken from "length" field
* **Width (in)**: taken from "width" field
* **Height (in)**: taken from "height" field
* **Weight (lb)**: taken from "productweight" field
* **Description**: taken from "productdescriptionshort" field
* **Image Link**: taken from "photourl" field
* **Product Link**: taken from "photourl" field
* **Categories**: list of categories joining "categoryids" with the category list
  
For access to a completed exe file, visit the "Release" section.

## How To Create Executable File

### Step 1: Download Product and Category CSV Files from Volusion
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
