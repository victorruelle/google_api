# Google API made friendly

Google's API are great but require quite some work to use efficiently. We've built some handy methods to:
- Setup tokens and authorizations in one command
- organize your drive with simple ```create_dir```, ```move``` and ```share``` methods
- create, load, read, write and style Google Sheets effortlesly
- create, load, write with styles Google Docs effortlesly

Work is based on the following API versions:
- drive : v3
- sheets : v4
- docs : v2

## Setup methods

To yourself up, run ```python setup.py``` with any of the ```-drive -docs -sheets``` options. This will automatically create credentials and tokens after taking you to Google's login page. Beware that you still need to enalbe the APIs on your account. To do so, you need to press the button "Enable ... API" for each API you mean to use:
- https://developers.google.com/drive/api/v3/quickstart/python
- https://developers.google.com/sheets/api/quickstart/python
- https://developers.google.com/docs/api/quickstart/python

## Drive APIs

1. Create a directory
``` python
from drive import create_dir

directory_id = create_dir("My folder")
subdirectory_id = create_dir("A subfolder",parent_id = directory_id)
```

2. Move a file
```python
from drive import move
move(file_id,directory_id)
```

3. Share a drive
```python
from drive import share
share(file_id,"my_partner@gmail.com",send_notification_email=False) # share a single file
share(directory_id,"my_partner@gmail.com",role="reader",send_notification_email=True) # share a drive but only with reading authorization
```
## Sheets APIs

1. Loading and creating a Sheet 
```python
from sheets import GoogleSheet

sheet = GoogleSheet.load(spreadsheetID)
sheet_new = GoogleSheet.new(sheet_name,optional parent directory ID)
```

2. Reading from a sheet
```python
sheet.read("A1:B2") # [['a1 cell','b1 cell'],['a2 cell','b2 cell']]
```

3. Writing to a sheet
```python
sheet.write("A6:B7",[["a6","b6"],['a7','b7']]) # modifies spreadsheet directly

sheet.add_write("A8:B9",[["a8","b8"],['a9','b9']])
sheet.add_write("C10",[['c10']])
sheet.execute_writes() # executes all writes at once
```

4. Styling a sheet
```python
cell_style_update = sheet.add_style_update("A3:B5")
cell_style_update.bold()
cell_style_update.fontsize(13)
cell_style_update.color(1,0.5,0)


cell_style_update_2 = sheet.add_style_update("A1")
cell_style_update_2.italic(False)
cell_style_update_2.underline()
cell_style_update_2.list_validation(["option 1","option 2"])


sheet.execute_style_updates() # executes all orders
```

Available methods are :
- bold : True or False
- italic : True or False
- underline : True or False
- strikethrough : True or False
- color : r,g,b
- background_color : r,g,b
- font : str (Font family)
- fontsize : float
- wrap : str (wrap statregy : WRAP, CELL_OVERFLOW, CLIP)
- border : str (direction : left,right,top,bottom or all), optional : style (DOTTED, SOLID, ...), color (r,g,b)
- list_validation : [str : values] 


## Docs APIs

1. Create and load a Doc
```python
from docs import GoogleDoc

new_doc = GoogleDoc.new("test document")
old_doc = GoogleDoc.load(DocumentID)
```

2. Write to a doc using formatting
```python

doc.add_write("Title\n","title")
doc.add_write("subtitle\n","subtitle")
doc.add_write("Header 1\n","heading 1")
doc.add_write("Content of heading 1\n")
doc.add_write("\n")
doc.add_write("Header 2\n","heading 2")
doc.add_write("Content of heading 2\n\n")
doc.execute()
```
Or write as a batch
```python

requests = [
    ("Title\n","title"),
    ("subtitle\n","subtitle"),
    ("Header 1\n","heading 1"),
    ("Content of heading 1\n"),
    ("\n"),
    ("Header 2\n","heading 2"),
    ("Content of heading 2...")
    ("...continue the current line.")
]

doc.add_writes(requests)
doc.execute()
```

## Next Steps

- [ ] Add reading and robust styling for GoogleDocs