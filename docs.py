import pickle
from googleapiclient.discovery import build
import json
from os import path

from drive import move_file

CURDIR = path.dirname(path.abspath(__file__))

class GoogleDoc():

    styled_types = {
        "title" : "TITLE",
        "heading 1" : "HEADING_1",
        "heading 2" : "HEADING_2",
        "heading 3" : "HEADING_3",
        "subtitle" : "SUBTITLE",
        "normal" : "NORMAL_TEXT"
    }

    def __init__(self,verbose=0):
        with open(path.join(CURDIR,"tokens",'token_docs.pickle'), 'rb') as token:
            creds = pickle.load(token)
        self.service = build('docs', 'v1', credentials=creds)
        self.documentId = None
        self.requests = []
        self.history = []
        self.verbose = verbose
        self.currentIndex = None
        self.url = 'https://docs.google.com/document/d/{}'

    def get_service(self):
        return self.service
    
    def execute(self):
        result = self.service.documents().batchUpdate(documentId=self.documentId, body={'requests': self.requests}).execute()
        if self.verbose:
            print(json.dumps(result,indent=2))
        self.history.append({"request":self.requests,"response":result})


    def view_history(self):
        print(json.dumps(self.history,indent=2))


    def add_write(self,text,style="normal"):
        self.insert(text,self.currentIndex,style)
        self.currentIndex += len(text)

    def add_writes(self,content):
        ''' Add requests to write content to the current document. By default, content will be added to the end of the document.
        For every piece of content, style can be specified. The list of possible styles is stored in GoogleDoc.styles_types.

        # Input
            - content : list of string (text) or tuples (text,style) or Content object. If no style is given (string), normal text will be applied.


        # Example
            document.write([
                ("My Title","title"),
                ("My subtitle","subtitle"),
                "Let's write some normal text",
                ("My first Header","Heading 1"),
                "This first header's content"
            ])
        '''
        for element in content:

            if type(element)==str:
                self.add_write(element)
            elif type(element) in [list,tuple]:
                self.add_write(element[0],element[1])
            elif type(element)==dict:
                self.add_write(element["text"],element["named_style"])
            else:
                raise Exception("Element type not supported : {} ({})".format(type(element),element))

        self.execute()


    def insert(self,text,index,style=None):
        request = {
            'insertText': {
                'location': {
                    'index': index,
                },
                'text': text
            }
        }
        self.requests.append(request)

        if style is not None or style=="normal":
            self.update_paragraph_style(style,index,index+len(text))


    def update_paragraph_style(self,style,startIndex,endIndex):
        request = {
            "updateParagraphStyle": {
                "paragraphStyle": {
                    "namedStyleType":GoogleDoc.styled_types[style]
                },
                "fields": "*",
                "range": {
                    "segmentId": "",
                    "startIndex": startIndex,
                    "endIndex": endIndex
                }
            }
        }
        self.requests.append(request)

    @classmethod
    def load(cls,documentId):
        doc = GoogleDoc()
        response = doc.service.documents().get(documentId=documentId).execute()
        doc.documentId = documentId
        doc.url = doc.url.format(doc.documentId)
        doc.currentIndex = response["body"]["content"][-1]["endIndex"]
        return doc

    @classmethod
    def new(cls,name,parent=None):
        doc = GoogleDoc()
        file_metadata = {
        'title': name
        }
        response = doc.service.documents().create(body=file_metadata).execute()
        if parent is not None:
            documentId = response.get("documentId")
            _ = move_file(documentId,parent)
        doc.documentId = response.get("documentId")
        doc.url = doc.url.format(doc.documentId)
        doc.history.append({"request":{"create":{"body":{"title":name}}},"response":response})
        doc.currentIndex = 1
        return doc

if __name__ == "__main__":
    doc = GoogleDoc.new("test1")
    doc._write("Title\n","title")
    doc._write("subtitle\n","subtitle")
    doc._write("Header 1\n","heading 1")
    doc._write("Content of heading 1\n")
    doc._write("\n")
    doc._write("Header 2\n","heading 2")
    doc._write("Content of heading 2\n\n")
    doc.execute()

