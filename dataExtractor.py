import cv2
import os,argparse
import pytesseract
from PIL import Image


def extractDataFromImage(imageFileName, imageBaseDir):
    try:
        PTPath = r"C:\\Users\\ASEN5\\AppData\\Local\\Programs\\Tesseract-OCR\\tesseract.exe"
        pytesseract.pytesseract.tesseract_cmd = PTPath
        text = pytesseract.image_to_string(Image.open(imageBaseDir + imageFileName))
        return text


    

    except Exception as Err:
        print("Error occruued while extrating the data from image", Err)

