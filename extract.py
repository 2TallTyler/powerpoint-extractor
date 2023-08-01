# We need to import these to fix an AttributeError
import collections
import collections.abc

# Actual imports
import os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import glob
import csv

_image_index = 0
_errors = 0

def save_image(image, name):
    '''
    Save an image to disk.

    :param image The image to save.
    :param name The partial name of the image, including the directory but not including the sequence number.
    '''
    global _image_index

    image_bytes = image.blob

    # Append the sequence number and file extension to the partial name
    name = name + ' ' + str(_image_index) + '.' + image.ext

    print(name)
    with open(name, 'wb') as f:
        f.write(image_bytes)

    _image_index += 1

def drill(shape, name):
    '''
    Recursive function to look inside grouped shapes for pictures and save them.

    :param shape The parent shape to look inside.
    :param name The partial name for any images we find, including the directory but not including the sequence number.
    '''
    global _image_index
    global _errors

    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for s in shape.shapes:
            drill(s, name)
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        try:
            save_image(shape.image, name)
        except:
            print("Could not process an image")
            _errors += 1

def iter_shapes(p):
    '''
    Iterate through shapes in the given Presentation

    :param p The presentation to iterate through.
    '''
    for slide in p.slides:
        for shape in slide.shapes:
            yield shape


def process_images(input_dir, output_dir):
    """
    Processes a folder of PowerPoint documents and saves each image to another folder.

    :param input_dir The name of the subfolder to search for PowerPoint documents.
    :param output_dir The name of the subfolder in which to save images.
    """
    global _image_index

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for eachfile in glob.glob(input_dir + os.sep + "*.pptx"):
        _image_index = 1

        # Strip input directory path and replace with chosen output path
        name = eachfile.split(os.sep)[1]

        # Strip file extension
        name = name.split('.')[0]

        # Prepend output directory to name, for saving later
        name = output_dir + os.sep + name

        for shape in iter_shapes(Presentation(eachfile)):
            drill(shape, name)

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

        print("Finished. Total files: " + str(count) + ", total errors reading images: " + str(_errors))

def main():
    input_dir = 'input'
    image_output_dir = 'images'

    process_images(input_dir, image_output_dir)
    process_text(input_dir)

if __name__ == "__main__":
    main()
