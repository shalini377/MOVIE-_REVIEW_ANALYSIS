#Importing Libraries
import speech_recognition as sr
from googletrans import Translator
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
import pandas as pd
import pyttsx3
from gtts import gTTS
import xlsxwriter
import openpyxl
import nltk
nltk.download('wordnet')
from nltk.corpus import wordnet
import os


engine=pyttsx3.init()
voices=engine.getProperty('voices')
engine.setProperty('rate', 150)
engine.setProperty('voice',voices[1].id)
def talk(text):
    engine.say(text)
    engine.runAndWait()
talk("mode slection")
talk("May i know what mode would you like to choose Admin or Reviewer:")
print("May i know what mode would you like to choose Admin or Reviewer:")

mode=str(input()).lower()

# movie storage

# def generate_excel(data:list):
#     workbook=xlsxwriter.Workbook('MOVIE.xlsx')
#     worksheet=workbook.add_worksheet('LIST.xlsx')

#     # adding headers
#     worksheet.write(0,0,"movie_name")
#     worksheet.write(0,1,"positive")
#     worksheet.write(0,2,"neutral")
#     worksheet.write(0,3,"negative")

#     for index,entry in enumerate(data):
#         worksheet.write(index+1,0,str(index))
#         worksheet.write(index+1,1,entry['movie_name'])
#         worksheet.write(index+1,2,entry['positive'])
#         worksheet.write(index+1,4,entry['neutral'])
#         worksheet.write(index+1,3,entry['negative'])
#     workbook.close()
excel_file = 'MOVIE.xlsx'

# Load the Excel file using pandas
df = pd.read_excel(excel_file)

# Specify the name of the column you want to extract
column_name = 'movie_name'

# Extract the column values and store them in a list
column_values = df[column_name].tolist()

def generate_excel(workbook_name: str, worksheet_name: str, headers_list: list, data: list):

    # Creating workbook
    workbook = xlsxwriter.Workbook(workbook_name)
    
    # Creating worksheet
    worksheet = workbook.add_worksheet(worksheet_name)

    # Adding headers
    for index, header in enumerate(headers_list):
        worksheet.write(0, index, str(header).capitalize())

    # Adding data
    for index1, entry in enumerate(data):
        for index2, header in enumerate(headers_list):
            worksheet.write(index1+1, index2, entry[header])
    # Close workbook
    workbook.close()
#movie list

if mode=="reviewer":
    print("Existing data",column_values)
    talk("movie name:")
    movie_name=str(input("movie name:")).upper()
    positive=negative=neutral=0


    #Microphone access
    # recognizer=sr.Recognizer()
    # with sr.Microphone() as source:
    #     print("Clearning the bac noise...")
    #     recognizer.adjust_for_ambient_noise(source, duration=2)
    #     print("Waiting for your message...")
    #     recordedaudio=recognizer.listen(source)
    #     print("Done Recording...")

    # # import whisper
    # # model = whisper.load_model("base")
    # # result = model.transcribe("output.mp3")
    # # print(result["text"])
    # try:
    #     print("Printing the message...")
    #     text=recognizer.recognize_google(recordedaudio,language="en-Us,tamil")
    #     print("Your message:{}".format(text))
    # except Exception as ex:
    #     print(ex)

    # text = text.lower()

    # #Storing the sentence
    # Sentence=[str(text)]
    # Voice = str(text)
    # file1 = open("record.txt", "a")  # append mode
    # file1.write("\n")
    # file1.write(Voice)
    # file1.close()

    # #google translator
    # k = Translator()
    # lang = "ta"
    # convert = "en"
    # output = k.translate(text,src=lang,dest=convert)
    # print(output)


    #review storage to excel
    def write_data_to_excel(file_name, sheet_name, data):
    # Create a new workbook or load an existing one
        try:
            wb = openpyxl.load_workbook(file_name)
        except FileNotFoundError:
            wb = openpyxl.Workbook()

        # Get the active sheet or create a new one
        sheet = wb[sheet_name] if sheet_name in wb else wb.create_sheet(sheet_name)

        # Append the new data set
        for row in data:
            sheet.append(row)

        # Save the excel file
        wb.save(file_name)
    def existing_data(name,out):
        file_path = 'MOVIE.xlsx'
        df = pd.read_excel(file_path)

        # Update the positive sentiment score for "Leo" to be incremented by 1
        df.loc[df['movie_name'] == name,out] += 1

        # Save the modified DataFrame back to the Excel file
        df.to_excel(file_path, index=False)

        print(f"{out} sentiment score for {name} has been updated.")


    #english dictionaray lemmatizer
    def is_english_word(word):
        synsets = wordnet.synsets(word)
        return len(synsets) > 0
    #dictonary mapping
    # text = output.text
    # text.lower()

    text="movie super macha mass a iruku"
    print(text)
    sample_data={"macha":"dude","mamey":"dude","semma":"bliss","maja":"fun","takkar":"joy","sumar":"impartial","mokka":"bad","tharama iruku":"super quality","tharu maru": "smash hit","attu mokka":"bore","kevalam":"awful","vera mari":"super hit","tharama":"quality","padam":"movie","veri a iruku":"super hit","mokka in maruuruvam":"monotonous","vera level":"out of box","nalla illa":"not good","palasu":"old","puthusa onum illa":"nothing new","padam":"movie","iruku":"have"}
    datas=""
    for i in text.split():
        if is_english_word(i):
            datas+=i+" "
        else:
            if i in sample_data:
                i=sample_data[i]
                datas+=i+" "
            else:
                datas+=i+""
    print(datas)
    #Sentimental Analyser
    output = [str(datas)]
    analyser = SentimentIntensityAnalyzer()
    for i in output:
        v=analyser.polarity_scores(i)
        print(v)
    
    out='value'
    #Output Detection
    if(v['compound']>0):
        out="positive"
        ans="Hey Your Message Sounds Positive.Hurray"
    elif(v['compound']==0):
        out="neutral"
        ans="Hey Your Message Sounds Neutral.Cool"
    else:
        out="negative"
        ans="Your Messages Sounds Sad!!! Dont Worry"

    print(ans)
    if out=="positive":
        positive+=1
    elif ans=="negative":
        negative+=1
    elif ans=="neutral":
        neutral+=1
    data=[movie_name,positive,negative,neutral]
    if movie_name not in column_values:
        write_data_to_excel('MOVIE.xlsx', 'Sheet1', [data])
    elif movie_name in column_values:
        print("Already existing data")
        existing_data(movie_name,out)

    #System Output
    talk("Do you want me to sound your review status?1")
    sound=str(input("Do you want me to sound your review status?(Y/N):"))
    sound.lower()
    if sound in["y","yes"]:
        mytext = ans
        language = 'en'
        output = gTTS(text=mytext, lang=language, slow=False)
        output.save("output.mp3")
        
        os.system("start output.mp3")
    else:
        print("thankyou for ur review!")
else:
# Read the Excel file
    file_name = 'MOVIE.xlsx'
    sheet_name = 'Sheet1'

    df = pd.read_excel(file_name, sheet_name=sheet_name)
    # Print the data set
    print(df)




