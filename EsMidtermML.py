# Brian Martin
# Es Midterm

# Imports the Google Cloud client library
from google.cloud import language_v1
import xlrd

# Instantiates a client
client = language_v1.LanguageServiceClient()

wb = xlrd.open_workbook(r"C:\Users\Brian\Documents\Wolfram Mathematica\Week 6\Lab 6\SDGIndicator_ML.xlsx")
ws = wb.sheet_by_index(0)
Indicator_list = ws.col_values(0)

wb = xlrd.open_workbook(r"C:\Users\Brian\Documents\Wolfram Mathematica\Week 6\Lab 6\SDGTarget_ML.xlsx")
ws = wb.sheet_by_index(0)
Target_list = ws.col_values(0)

Indicator_Dictionary = {}
Target_Dictionary = {}
def Setiment_Analysis(Text_List,type):
    # The text to analyze
    n=1
    for text in Text_List:
        document = language_v1.Document(content=text, type_=language_v1.Document.Type.PLAIN_TEXT)

        # Detects the sentiment of the text
        sentiment = client.analyze_sentiment(request={'document': document}).document_sentiment

        key = type + str(n)
        if type == "Indicator_":
            Indicator_Dictionary[key] = [sentiment.score, sentiment.magnitude]
        elif type == "Target_":
            Target_Dictionary[key] = [sentiment.score, sentiment.magnitude]

        # print("Text: {}".format(text))
        # print("Sentiment: {}, {}".format(sentiment.score, sentiment.magnitude))
        n+= 1


Setiment_Analysis(Indicator_list,"Indicator_")
Setiment_Analysis(Target_list,"Target_")
print(Indicator_Dictionary.keys())
print(Indicator_Dictionary.values())
print(len(Indicator_Dictionary))

print(Target_Dictionary.keys())
print(Target_Dictionary.values())
print(len(Target_Dictionary))