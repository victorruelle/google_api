from spreadsheets import GoogleSheet

sheet = GoogleSheet.load()

# sheet.write("B5",[["BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB"]])

sheet.add_style_update("B5").list_validation(["1","2"])

sheet.execute_style_updates()
