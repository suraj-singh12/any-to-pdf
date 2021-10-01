import os
import sys

import fpdf     # for txt to pdf
from win32com import client     # for excel to pdf
from docx2pdf import convert
from ppt2pdf import main as ppt_processor
import img2pdf
from PIL import Image
from io import BytesIO


def generateOutputFilename(filename):
    filename = os.path.join(os.getcwd(), filename)
    output = os.path.splitext(filename)
    return os.path.abspath(output[0]+".pdf")


def Text2Pdf(filename):
    pdf = fpdf.FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=10)

    f = open(filename, 'r')
    # writing content to pdf

    for line in f:         # for each line
        # split the line into list of words
        listOfWords = list(line.split())

        indx = 0
        upperBnd = 15
        while indx + upperBnd < len(listOfWords):
            # itr1: first 15 words, itr2: next 15 words, itr3: next 15 words and so on..
            string = ' '.join(listOfWords[indx: indx + upperBnd])
            pdf.cell(200, 10, txt=string, ln=1, align='L')

            # then update the start index to reach next 15 words in another iteration
            indx += upperBnd

        # if some words are left, then write them too
        if len(listOfWords) - indx > 1:
            string = ' '.join(listOfWords[indx: len(listOfWords)])
            print(string)
            pdf.cell(200, 10, txt=string, ln=1, align='L')

    # pdf file saved
    output_filename = generateOutputFilename(filename)
    pdf.output(output_filename)
    print("[+] Saved at:", output_filename)


def Ppt2Pdf(filename):
    # executing the command, conversion done !!
    output_filename = generateOutputFilename(filename)
    ppt_processor.convert(filename, output_filename)
    print("[+] Saved at:", output_filename)


def Excel2Pdf(filename):
    # open microsoft excel (using it's built-in functionality)
    excel = client.Dispatch("Excel.Application")

    # reading excel file
    sheets = excel.Workbooks.open(os.path.abspath(filename))
    work_sheets = sheets.Worksheets[0]

    # creating new filename with .pdf extension to save resulting file in current dir
    new_file_name = generateOutputFilename(filename)
    # converting to pdf file
    work_sheets.ExportAsFixedFormat(0, os.path.abspath(new_file_name))
    print("[+] Saved at:", new_file_name)


def Docx2Pdf(filename):
    # conversion done
    convert(filename)


def Image2Pdf(filename):
    # open image in PIL
    im = Image.open(filename)
    # checks for transparency
    if (
        im.mode in ('RGBA', 'LA') or
        (im.mode == 'P' and 'transparency' in im.info)
    ):
        im.load()
        background = Image.new("RGB", im.size, (255, 255, 255))
        background.paste(im, mask=im.split()[3])
        im = background
    output_filename = generateOutputFilename(filename)
    im.save(output_filename)
    print("[+] Saved at:", output_filename)


def main():
    print("---->Welcome To Universal PDF Converter<----\n")
    filename = input("Enter the file name(with extension): ")

    if not os.path.exists(filename):
        print("Error! no such file present in current directory.")
        sys.exit(-1)
    if filename.endswith(".txt"):
        Text2Pdf(filename)
    elif filename.endswith(".ppt") or filename.endswith(".pptx"):
        Ppt2Pdf(filename)
    elif filename.endswith(".docx"):
        Docx2Pdf(filename)
    elif filename.endswith(".xlsx"):
        Excel2Pdf(filename)
    elif any(filename.endswith(ext) for ext in ["jpg", "png", "bmp"]):
        Image2Pdf(filename)


if __name__ == "__main__":
    main()
