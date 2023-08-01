# PowerPoint Extractor
A simple script to extract images, text, and presenter notes from a folder full of PowerPoint files. It uses the `python-pptx` library.

## Instructions
You need [python-pptx](https://pypi.org/project/python-pptx/), if you don't have it already. Install with:
```pip install python-pptx```

1. Clone the repository onto your local drive.
2. Copy PowerPoint files into the [input](/input) folder.
3. Run `extract.py`.

## Output
* Text will be saved to a new `text.csv` file in the root folder. This has a row for each slide, with columns containing the presentation name, page number, all the text from the page, and any presenter notes.
* Images will be saved to a new `images` folder, named sequentially with the name of the presentation.

## Development
This is a quick and dirty script I wrote for a specific project. I welcome PRs to clean up code, add features, etc.
