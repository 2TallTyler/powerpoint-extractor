# We need to import these to fix an AttributeError
import collections
import collections.abc

# Actual imports
import os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import glob
import csv

def iter_picture_shapes(prs):
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                yield shape

def process_images(input_dir, output_dir):
    """
    Processes a folder of PowerPoint documents and saves each image to another folder.

    :param input_dir The name of the subfolder to search for PowerPoint documents.
    :param output_dir The name of the subfolder in which to save images.
    """

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for eachfile in glob.glob(input_dir + os.sep + "*.pptx"):
        image_number = 1

        for picture in iter_picture_shapes(Presentation(eachfile)):
            image = picture.image
            image_bytes = image.blob

            # Strip directory path and replace with chosen output path
            image_filename = eachfile.split(os.sep)[1]
            image_filename = output_dir + os.sep + image_filename

            # Name the image sequentially
            image_filename = image_filename + ' ' + str(image_number) + '.' + image.ext

            print(image_filename)
            with open(image_filename, 'wb') as f:
                f.write(image_bytes)

            image_number += 1

def process_text(input_dir):
    """
    Processes a folder of PowerPoint documents and saves a .csv report of all the text within each document, including presenter notes.

    :param input_dir The name of the subfolder to search for PowerPoint documents.
    """
    with open('text.csv', 'w', encoding="utf-8", newline='') as file:
        writer = csv.writer(file)
        count = 0

        # Write table header
        field = ["file", "page", "text", "notes"]
        writer.writerow(field)

        # Iterate through files
        print("Processing:")
        for eachfile in glob.glob(input_dir + os.sep + "*.pptx"):
            ppt = Presentation(eachfile)
            print("* " + eachfile)
            count += 1

            # Iterate through slides
            for page, slide in enumerate(ppt.slides):
                text = ''

                # Collect all the text on the slide into one string, separated by newlines
                for shape in slide.shapes:
                    if shape.has_text_frame and shape.text.strip():
                        text += os.linesep
                        text += shape.text

                # Write the page number, collected text, and presenter notes as a new row
                writer.writerow([eachfile, page, text, slide.notes_slide.notes_text_frame.text])

        print("Finished. Total files: " + str(count))

def main():
    input_dir = 'input'
    image_output_dir = 'images'

    process_images(input_dir, image_output_dir)
    process_text(input_dir)

if __name__ == "__main__":
    main()
