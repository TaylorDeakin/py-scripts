import zipfile
from os import listdir, getcwd
from shutil import rmtree
import xml.etree.cElementTree
import sys

if len(sys.argv) is not 2:
    print("usage: ppt-extract.py <filename>")
    exit()

# open the powerpoint as a zip - you can legit do this, it's awesome
with zipfile.ZipFile(sys.argv[1]) as z:
    # we need to extract the zip file to get access to all the things
    z.extractall("out")

# since we've extracted the ppt to the working directory, we use getcwd, then append the path to the slides directory
slide_folder = getcwd() + "\\out\\ppt\\slides"

# from there, we can get all the slides as .xml files
slides = listdir(slide_folder)
# there's also a directory called _rels, which we don't want
# luckily it's always the last item in the list
slides.pop()

# unfortunately, we can't guarantee the order of the list at this point
# so we sort it with a lambda function
# the function takes a string with the form slideX.xml and returns X as an int
# this can then be used for sorting purposes
slides.sort(key=lambda str: int((str.split(".")[0]).lstrip("slide")))


# parse_slide takes a slide and returns an array of text from it
def parse_slide(slide):
    slide_text = []
    # utf8 encoding is important!
    with open(slide_folder + "\\" + slide, encoding='utf8') as f:
        tree = xml.etree.cElementTree.parse(f)
        # iterate over the text, and get the specified tag
        for tag in tree.iter(tag='{http://schemas.openxmlformats.org/drawingml/2006/main}p', ):
            text_bits = ""
            for text_tag in tag.iter(tag='{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
                if text_tag.text is not None:
                    text_bits += text_tag.text + " "

            slide_text.append(text_bits)
    f.close()
    return slide_text


# get the text from the slides
slides_text = []
for slide in slides:
    slides_text.append(parse_slide(slide))

filename = (sys.argv[1].split("."))[0]

# save it to a file
with open(filename + ".txt", encoding='utf8', mode='w') as outfile:
    for slide in slides_text:
        for line in slide:
            if line is not None:
                outfile.write(line + "\n")

    outfile.close()

# finally, we need to tidy up after ourselves, so we remove the folder
rmtree(getcwd() + "\\out")
