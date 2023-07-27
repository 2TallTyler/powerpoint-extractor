# PowerPoint Extractor
A simple script to extract images, text, and presenter notes from a folder full of PowerPoint files. It uses the `python-pptx` library.

## Instructions
You need [python-pptx](https://pypi.org/project/python-pptx/), if you don't have it already. Install with:
```pip install python-pptx```

1. Clone the repository onto your local drive.
2. Copy PowerPoint files into the [input] folder.
3. Run `extract.py`.

The script will create a new `text.csv` file containing all the text from each slide, alongside the presentation name, page number, and presenter notes.

Images will be saved to a new `images` folder, named sequentially with the name of the presentation.

## To do
This is a quick and dirty script I wrote for my job. Feel free to open PRs to clean up my code, add features, etc.

A few things I might want to add in the future:
* Add the slide number to the image filename
* Split into separate scripts to only process images or text, instead of both
