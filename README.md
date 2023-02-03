# PresentationBuilder

> ## This is a Python module for building Powerpoint presentations using the pptx library. The main aim of this module is to create presentations for a worship service, with dynamic Bible Portions and Song Lyrics.

## Requirements

1.  [Python 3.8+](https://www.python.org/downloads/)
2.  These fonts must be installed on your system:

    -   [Noto Serif Malayalam](https://fonts.google.com/noto/specimen/Noto+Serif+Malayalam?query=noto+serif+mala)
    -   [Goudy Bookletter 1911](https://fonts.google.com/specimen/Goudy+Bookletter+1911?query=goudy)

## Usage

-   ### Clone the repository (or just download)

    ` git clone https://github.com/gijocode/Church_PPT_Utility`

-   ### Install the required libraries

    `pip install -r requirements.txt`

-   ### Run PresentationBuilder.py and follow the prompts

    `python3 PresentationBuilder.py`

-   ### The output file will be saved in the same directory as Updated_PPT.pptx

---

## Structure

The code is structured as follows:

    1. PresentationBuilder class contains all the methods for building the presentation
    2. BibleExtractor is a helper class for extracting Bible portions
    3. PresentationBuilder.py is the entry point for running the code
    4. templates directory contains the template for the presentation
    5. assets directory contains additional assets used in the presentation

## Limitations

-   This project took _(annoyingly)_ more time than expected because of the severe limitations of the `python-pptx` library. If you examine the code you will see that even for obvious operations (such as adding a slide at a particular location in the Presentation), hacky workarounds needed to be used. But since this is the only library used for python + ppt, had to use the same.
-   The songs data used only covers the songs from the **Kristeeya Keertanangal** book, so additional songs need to be manually added

## Future Enhancements

-   Add support for English Service PPT
-   Add multicolor support for the verses slides, Malayalam - black text and English - blue text
-   GUI

## Contributions

Contributions to this project are welcome! Feel free to submit pull requests for bug fixes or new features.
