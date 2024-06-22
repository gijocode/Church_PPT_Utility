## Church Service Presentation Builder

This is a Python script for automating tedious presentation preparation for Mar Thoma Church services. It is designed to work with Microsoft PowerPoint files (.pptx) and provides a simple and intuitive command-line interface for building the slides.

### Features

- Automatically generates church service presentations based on a custom template.
- Integrates Bible verses, song lyrics, and other service elements into the slides.
- Easy-to-use command-line interface for adding content to the presentation.
- Supports both English and Malayalam church services.
- Provides options for customizing the theme and content for each service.

### Requirements

Before running the script, make sure you have the following prerequisites:

1.  [Python 3.8+](https://www.python.org/downloads/)
2.  These fonts must be installed on your system:

    - [Noto Serif Malayalam](https://fonts.google.com/noto/specimen/Noto+Serif+Malayalam?query=noto+serif+mala)
    - [Goudy Bookletter 1911](https://fonts.google.com/specimen/Goudy+Bookletter+1911?query=goudy)

### How to Use

1. Clone or download the repository to your local machine.

2. Open a terminal or command prompt in the project directory.

3. Install the required libraries

```bash
  pip install -r requirements.txt
```

4. Run PresentationBuilder.py and follow the prompts

```bash
  python3 PresentationBuilder.py
```

5. Follow the on-screen instructions to enter the service details, Bible portions, song numbers, and other content for the presentation.

6. Once you've provided all the required information, the script will generate the presentation and save it as a PowerPoint (.pptx) file.

7. The generated presentation will be saved in the project directory on Windows, or in the specified path on macOS.

### Adding Custom Templates

You can create your own custom templates for different types of church services. To do this, follow these steps:

1. Create a new PowerPoint file (.pptx) and design your custom slides using PowerPoint's features.

2. Save the file in the `templates` folder with a unique name that represents the type of service (e.g., `malayalam_service_template.pptx`, `english_service_template.pptx`).

3. Ensure that each slide contains a unique shape name for easy identification in the script.

4. Update the `REQ_SLIDES` list in the script with the shape names of the slides you want to populate dynamically. This list defines the order of the slides in the presentation.

5. Run the script and choose the appropriate service type to use your custom template.

### Note

1. Currently, I was able to only source lyrics for Kristeeya Geethangal. Therefore only songs from that book will have lyrics shown.
2. I highly encourage you to install [fzf](https://github.com/junegunn/fzf). This greatly improves the bible portion selection user interaction.

Feel free to contribute to this project and share your feedback!

**Happy presenting!**
