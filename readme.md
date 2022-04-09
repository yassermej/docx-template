# Generate docx files from template

## Set up environments
```
pip install -r requirements.txt
```

## Make sure that input file is valid json format
Remove comma at the end of last items in arry and dict. Check data_text.txt

## Run script
```
python generate.py
```

## Structure
1. data_text.xt # input json file
2. logo.png # logo file
3. template.docx # template docx file

## Modified Fields
1. I added `{{ suitability_table_caption }}` and `{{ suitability_criteria_table_caption }}` for table captions