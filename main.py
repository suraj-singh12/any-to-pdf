import os
import sys
from typing import Optional

import fpdf  # for txt to pdf
from win32com import client  # for excel to pdf
from docx2pdf import convert
from ppt2pdf import main as ppt_processor
from PIL import Image


def generateOutputFilename(filename: str, dest: str) -> str:
    filename = os.path.join(dest, os.path.split(filename)[-1])
    output = os.path.splitext(filename)
    return os.path.abspath(output[0] + ".pdf")


def Text2Pdf(filename: str, dest: str) -> str:
    pdf = fpdf.FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=10)

    f = open(filename, "r")
    # writing content to pdf

    for line in f:  # for each line
        # split the line into list of words
        listOfWords = list(line.split())

        indx = 0
        upperBnd = 15
        while indx + upperBnd < len(listOfWords):
            # itr1: first 15 words, itr2: next 15 words, itr3: next 15 words and so on..
            string = " ".join(listOfWords[indx : indx + upperBnd])
            pdf.cell(200, 10, txt=string, ln=1, align="L")

            # then update the start index to reach next 15 words in another iteration
            indx += upperBnd

        # if some words are left, then write them too
        if len(listOfWords) - indx > 1:
            string = " ".join(listOfWords[indx : len(listOfWords)])
            pdf.cell(200, 10, txt=string, ln=1, align="L")

    # pdf file saved
    output_filename = generateOutputFilename(filename, dest)
    pdf.output(output_filename)
    return output_filename


def Ppt2Pdf(filename: str, dest: str) -> str:
    # executing the command, conversion done !!
    output_filename = generateOutputFilename(filename, dest)
    ppt_processor.convert(filename, output_filename)
    return output_filename


def Excel2Pdf(filename: str, dest: str) -> str:
    # open microsoft excel (using it's built-in functionality)
    excel = client.Dispatch("Excel.Application")

    # reading excel file
    sheets = excel.Workbooks.open(filename)
    work_sheets = sheets.Worksheets[0]

    # creating new filename with .pdf extension to save resulting file in current dir
    new_file_name = generateOutputFilename(filename, dest)
    # converting to pdf file
    work_sheets.ExportAsFixedFormat(0, new_file_name)
    return new_file_name


def Docx2Pdf(filename: str, dest: str) -> str:
    # conversion done
    convert(filename, dest)
    return generateOutputFilename(filename, dest)


def Image2Pdf(filename: str, dest: str) -> str:
    # open image in PIL
    im = Image.open(filename)
    # checks for transparency
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        im.load()
        background = Image.new("RGB", im.size, (255, 255, 255))
        background.paste(im, mask=im.split()[3])
        im = background

    output_filename = generateOutputFilename(filename, dest)
    im.save(output_filename)
    return output_filename


def main(filename: str, dest: str = os.getcwd()) -> str:
    filename = os.path.abspath(filename)

    if not os.path.exists(filename):
        print("Error! no such file present in current directory.", file=sys.stderr)
        exit(-1)
    
    if filename.endswith(".txt"):
        return Text2Pdf(filename, dest)
    elif filename.endswith(".ppt") or filename.endswith(".pptx"):
        return Ppt2Pdf(filename, dest)
    elif filename.endswith(".docx"):
        return Docx2Pdf(filename, dest)
    elif filename.endswith(".xlsx"):
        return Excel2Pdf(filename, dest)
    elif any(filename.endswith(ext) for ext in ["jpg", "png", "bmp"]):
        return Image2Pdf(filename, dest)


if __name__ == "__main__":
    print("---->Welcome To Universal PDF Converter<----\n")
    out = main(input("Enter the file name (with extension): "))
    print(f"[+] Saved at: {out}")