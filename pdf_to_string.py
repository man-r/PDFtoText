# downloaded python 3.9.1 from https://www.python.org/downloads/  
# add the path C:\Users\[username]\AppData\Local\Programs\Python\Python39
# add the path C:\Users\[username]\AppData\Local\Programs\Python\Python39\Scripts
# install tesseract-ocr-w64-setup-v5.0.0-alpha.20200328.exe
# add teseract to path C:\Program Files\Tesseract-OCR
# install ImageMagick-7.0.10-35-Q16-HDRI-x64-dll.exe
# add imagemagic to path C:\Program Files\ImageMagick-7.0.10-Q16-HDRI
# install gs9533w64.exe
# pip install wand
# pip install Pillow
# pip install pytesseract
# pip install pandas
# pip install xlrd==1.2.0

# disabled maxlength path (if needed)


import os
import pandas as pd
import io
from wand.image import Image as wi
from PIL import Image
import pytesseract

def get_CV_score(CVFileName, keywordsFileName):
    data = pd.read_excel (keywordsFileName)                     # keywords into pandas data frame
    df = pd.DataFrame(data)
    
    pdfFilename = CVFileName                                    # read CV file
    pdf = wi(filename=pdfFilename, resolution=300)              # image setting
    pdfImage = pdf.convert("jpeg") 

    i = 1
    imgBlobs = []
    for img in pdfImage.sequence:                               # convert each page into an image and store the blob into an array
        page = wi(image=img)
        imgBlobs.append(page.make_blob('jpeg'))
        #page.save(filename=str(i)+".jpg")
        i += 1

    extracted_text = []
    for imgBlob in imgBlobs:                                    # ocr each image and save the text into an array
        im = Image.open(io.BytesIO(imgBlob))
        text = pytesseract.image_to_string(im, lang="eng")
        # print(text.encode("utf-8"))
        extracted_text.append(text)
        # delete the image

    score = 0
    found_keywords= {"manar"}                                   # initializing the found key words list
    for i, val in enumerate(extracted_text):                    # loop through every page text
        for index, row in df.iterrows():                        # loop through every keywords
            inSet = row[0] in found_keywords                    # check if the keyword has been found in previoce pages
            if (val.find(row[0]) > -1) and (inSet == False) :   # check if the keyword has been found for the first time
                found_keywords.add(row[0])                      # add the keyword to the found keywords list
                score = score + row[1]                          # update the CV score

    print(pdfFilename, " score = ", score, " keywords found: ", found_keywords) # print the final result
    allScores.append([pdfFilename, score, found_keywords]) # add the final result to the allSores list




# code starts here
allScores = [] # this ware all the file scores will be saved
arr = os.listdir('.') # get list of all file in the direcorey
for index in range(len(arr)): 
    print('file name', arr[index])


for index in range(len(arr)): 
    if arr[index].endswith('pdf'): # get the score for PDF only
        get_CV_score(arr[index], r'Audit Analytics Keywords(37) - Copy.xlsx') # key words excel file is hardcoded. the excel file must have two colums 1st contins keyword 2nd contains the score

print(allScores)
allScores_df = pd.DataFrame(allScores, columns = ['CV', 'Score', 'keywords']) # save the list into a dataframe cuz its more awesome!
print(allScores_df)
allScores_df.to_csv(r'exportallScores_df.csv', index=False) # export to csv