import pickle
from googleapiclient.discovery import build
import json
from os import path
import string

from drive import move_file
CURDIR = path.dirname(path.abspath(__file__))


class GoogleSheet():

    known_documents = {
        "testing_document" : "1vEfj-yXMLamKmlvX7jvy-cbyj36xUqMa9hNkaGEJGRI"
    }

    input_options = ["RAW","USER_ENTERED"]

    def __init__(self,verbose=0):
        '''
        '''
        with open(path.join(CURDIR,"tokens",'token_docs.pickle'), 'rb') as token:
            creds = pickle.load(token)
        self.service = build('sheets', 'v4', credentials=creds)
        self.spreadsheetID = None
        self.write_requests = []
        self.style_requests = []
        self.history = []
        self.verbose = verbose
        self.url = 'https://docs.google.com/document/d/{}'

    def get_service(self):
        return self.service
    
    def execute(self):
        result = self.service.spreadsheets().batchUpdate(spreadsheetID=self.spreadsheetID, body={'requests': self.requests}).execute()
        if self.verbose:
            print(json.dumps(result,indent=2))
        self.history.append({"request":self.requests,"response":result})


    def view_history(self):
        print(json.dumps(self.history,indent=2))

    
    def read(self,range_name):
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheetID, range=range_name).execute()
        rows = result.get('values', [])
        return rows

    def read_multiple(self,range_name):
        result = self.service.spreadsheets().values().batchGet(
            spreadsheetId=self.spreadsheetID, ranges=range_name).execute()
        rows = result.get('valueRanges', [])
        return rows

    def write(self,range_name,values):      
        body = {
            'values': values
        }
        result = self.service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheetID, range=range_name,
            valueInputOption=GoogleSheet.input_options[1], body=body).execute()

    def add_write(self,range_name,values):
        self.write_requests.append({"range":range_name,"values":values})

    def execute_writes(self):
        body = {
            'valueInputOption': GoogleSheet.input_options[1],
            'data': self.write_requests
        }
        result = self.service.spreadsheets().values().batchUpdate(
            spreadsheetId=self.spreadsheetID, body=body).execute()

        self.write_requests = []

    def add_style_update(self,range_name,sheetIndex=0):
        ''' Returns a CellStyleObject. Call it's methods to modify fields as desired. Finish by call execute_updates on self.
        '''
        cell_style_update = CellStyleUpdate(range_name,sheetIndex)
        self.style_requests.append(cell_style_update)
        return cell_style_update

    def execute_style_updates(self):

        body = {
            'requests': [req.to_json() for req in self.style_requests]
        }

        print("Body of request is",body)

        response = self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheetID,
            body=body).execute()
        self.style_requests = []
        print(response)



    @classmethod
    def load(cls,spreadsheetID=None):
        spreadsheetID = spreadsheetID or GoogleSheet.known_documents["testing_document"]
        sheet = GoogleSheet()
        response = sheet.service.spreadsheets().get(spreadsheetId=spreadsheetID).execute()
        sheet.spreadsheetID = spreadsheetID
        sheet.url = sheet.url.format(sheet.spreadsheetID)
        return sheet

    @classmethod
    def new(cls,name,parent=None):
        sheet = GoogleSheet()
        file_metadata = {
            'properties':{
                'title': name
            }
        }
        response = sheet.service.spreadsheets().create(body=file_metadata,fields='spreadsheetId').execute()
        if parent is not None:
            spreadsheetID = response.get("spreadsheetId")
            _ = move_file(spreadsheetID,parent)
        sheet.spreadsheetID = response.get("spreadsheetId")
        sheet.url = sheet.url.format(sheet.spreadsheetID)
        sheet.history.append({"request":{"create":{"body":{"title":name}}},"response":response})
        return sheet


class CellStyleUpdate():

    def __init__(self,range_name,sheetIndex=0):
        self.cell_style = CellStyle()
        startRowIndex,endRowIndex,startColumnIndex,endColumnIndex = decode_range(range_name)
        self.nrows,self.ncols = endRowIndex - startRowIndex, endColumnIndex - startColumnIndex
        self.range = {
            "sheetId": sheetIndex,
            "startRowIndex": startRowIndex,
            "endRowIndex": endRowIndex,
            "startColumnIndex": startColumnIndex,
            "endColumnIndex": endColumnIndex
        }
        self.fields = ""

    def add_field(self,path):
        self.fields += ("," if len(self.fields)>0 else "")+path

    def __repr__(self):
        return str(self.to_json())

    def to_json(self):
        return {
            "updateCells":{
                "rows":[ { "values": [self.cell_style.to_json() for _ in range(self.ncols)] } for _ in range(self.nrows) ],
                "fields":self.fields,
                "range":self.range
            }
        }

    def color(self,r,g,b,a=0):
        path = self.cell_style.color(r,g,b,a)
        self.add_field(path)
        return self

    def bold(self,state=True):
        path = self.cell_style.bold(state)
        self.add_field(path)
        return self

    def italic(self,state=True):
        path = self.cell_style.italic(state)
        self.add_field(path)
        return self

    def underline(self,state=True):
        path = self.cell_style.underline(state)
        self.add_field(path)
        return self

    def strikethrough(self,state=True):
        path = self.cell_style.strikethrough(state)
        self.add_field(path)
        return self

    def font(self,fontFamily):
        path = self.cell_style.font(fontFamily)
        self.add_field(path)
        return self

    def fontsize(self,size):
        path = self.cell_style.fontsize(size)
        self.add_field(path)
        return self

    def background_color(self,r,g,b,a=0):
        path = self.cell_style.background_color(r,g,b,a)
        self.add_field(path)
        return self

    def border(self,direction,style="SOLID",r=0,g=0,b=0,a=0):
        ''' Style is one of : DOTTED, DASHED, SOLID, SOLID_MEDIUM, SOLID_THICK, NONE, DOUBLE
        '''
        path = self.cell_style.border(direction,style,r,g,b,a)
        self.add_field(path)
        return self

    def wrap(self,stratgy="WRAP"):
        ''' Strategy is one of : WRAP, CLIP, OVERFLOW_CELL
        '''
        path = self.cell_style.wrap(stratgy)
        self.add_field(path)
        return self
    
    def list_validation(self,values,strict=True,showCustomUi=True):
        path = self.cell_style.list_validation(values,strict,showCustomUi)
        self.add_field(path)
        return self


def decode_range(range_name):
    parts = range_name.split(":")

    if len(parts)==1:
        code = parts[0]
        cut = find_cut(code)
        letters,numbers = code[:cut].lower(),int(code[cut:])-1
        startColumnIndex = 26*(len(letters)-1)+string.ascii_lowercase.index(letters[-1])
        endColumnIndex = startColumnIndex+1
        startRowIndex = numbers
        endRowIndex = startRowIndex+1

    elif len(parts)==2:

        # start
        code = parts[0]
        cut = find_cut(code)
        letters,numbers = code[:cut].lower(),int(code[cut:])-1
        startColumnIndex = 26*(len(letters)-1)+string.ascii_lowercase.index(letters[-1])
        startRowIndex = numbers

        # end
        code = parts[1]
        cut = find_cut(code)
        letters,numbers = code[:cut].lower(),int(code[cut:])-1
        endColumnIndex = 26*(len(letters)-1)+string.ascii_lowercase.index(letters[-1])+1
        endRowIndex = numbers+1
    else:
        raise Exception("Wrong format :{}".format(range_name))

    return startRowIndex,endRowIndex,startColumnIndex,endColumnIndex

def find_cut(code):
    cut = 0
    while not str.isnumeric(code[cut]):
        try:
            cut += 1
        except IndexError:
            print("Code is not interpretable : {}".format(code))
    return cut
        

class CellStyle():

    def __init__(self):
        self.body = Property()

    def to_json(self):
        return self.body.to_json()

    def color(self,r,g,b,a=0):
        self.body.to("userEnteredFormat").to("textFormat").to("foregroundColor").update({"red":r,"green":g,"blue":b,"alpha":a})
        return "userEnteredFormat.textFormat.foregroundColor"

    def bold(self,state=True):
        self.body.to("userEnteredFormat").to("textFormat").update({"bold":state})
        return "userEnteredFormat.textFormat.bold"

    def italic(self,state=True):
        self.body.to("userEnteredFormat").to("textFormat").update({"italic":state})
        return "userEnteredFormat.textFormat.italic"

    def underline(self,state=True):
        self.body.to("userEnteredFormat").to("textFormat").update({"underline":state})
        return "userEnteredFormat.textFormat.underline"

    def strikethrough(self,state=True):
        self.body.to("userEnteredFormat").to("textFormat").update({"strikethrough":state})
        return "userEnteredFormat.textFormat.strikethrough"

    def font(self,fontFamily):
        self.body.to("userEnteredFormat").to("textFormat").update({"fontFamily":fontFamily})
        return "userEnteredFormat.textFormat.fontFamily"

    def fontsize(self,size):
        self.body.to("userEnteredFormat").to("textFormat").update({"fontSize":size})
        return "userEnteredFormat.textFormat.fontSize"

    def background_color(self,r,g,b,a=0):
        self.body.to("userEnteredFormat").to("backgroundColor").update({"red":r,"green":g,"blue":b,"alpha":a})
        return "userEnteredFormat.backgroundColor"

    def border(self,direction,style="SOLID",r=0,g=0,b=0,a=0):
        ''' 
        # Input
            - Direction : left,right,top,bottom or all
            - Style : DOTTED, DASHED, SOLID, SOLID_MEDIUM, SOLID_THICK, NONE, DOUBLE
        '''
        if direction == "all":
            for direction in ["left","right","bottom","top"]:
                self.body.to("userEnteredFormat").to("borders").to(direction).update({"style":style})
                self.body.to("userEnteredFormat").to("borders").to(direction).to("color").update({"red":r,"green":g,"blue":b,"alpha":a})

        else:
            self.body.to("userEnteredFormat").to("borders").to(direction).update({"style":style})
            self.body.to("userEnteredFormat").to("borders").to(direction).to("color").update({"red":r,"green":g,"blue":b,"alpha":a})
        
        return "userEnteredFormat.borders"

    def wrap(self,stratgy="WRAP"):
        ''' Strategy is one of : WRAP, CLIP, OVERFLOW_CELL
        '''
        self.body.to("userEnteredFormat").update({"wrapStrategy":stratgy})
        return "userEnteredFormat.wrapStrategy"

    def list_validation(self,values,strict=True,showCustomUi=True):
        self.body.to("dataValidation").update({
            "condition":{
                "type":"ONE_OF_LIST",
                "values": [
                    {"userEnteredValue":value}
                for value in values]
            },
            "strict":strict,
            "showCustomUi":showCustomUi
        })
        return "dataValidation"

    def __repr__(self):
        return str(self.body)



class Property():

    def __init__(self):
        self.body = {}

    def to(self,key):
        if key in self.body:
            assert isinstance(self.body[key],Property),"Use to only to go through properties, key : {}, body : {}".format(key,self.body)
            return self.body[key]
        else:
            self.body[key] = Property()
            return self.to(key)

    def set(self,key,value):
        self.body[key]=value

    def update(self,d):
        self.body.update(d)

    def to_json(self):
        return {key:(val.to_json() if isinstance(val,Property) else val) for key,val in self.body.items()}

    def __repr__(self):
        return str(self.body)