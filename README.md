# Excel Sheet Collator 
Automatically collate data tables in Excel workbooks. 

## How it works 
Collates data tables of separate Excel workbooks into a single table using common identifiers. 
Demonstrated in the image below:

<img alt="GUI" src="https://i.imgur.com/BTVzvHg.png.png" width="90%"></img>

## Getting started 
Download: [Python](https://www.python.org/downloads/), [Pillow](https://pypi.org/project/Pillow/), [Pandas](https://pypi.org/project/pandas/) 

## Running the application 
<img alt="GUI" src="https://i.imgur.com/mv5bVtz.png" width="30%"></img>

- Ensure Excel sheets are of type ```.xlsx```
- Excel sheets should reside in the same directory as the application
- Specify the **’index name’** in the input bar. 
    - This is the column of unique identifiers. 
    - **Example**: In the previous image the ‘index name’ would be the column ‘ID’. 
  - **Note**: This is case and character sensitive! 
- Click **‘Collate Excel Sheets`**
