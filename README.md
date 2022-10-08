# Church_PPT_Utility

A simple python utility that helps populate a church PPT with the requireed Lesson Portions, Song info etc.

## Prerequisites

This utility requires python3. If you do not have it installed on your system, you can install it from: https://www.python.org/downloads/
Also, add python3 to your PATH variable

This project works only on the PPT in the `templates` folder. You can edit the PPT as long as you don't edit/delete the song/lesson slides and it should (hopefully) work fine.

## Usage

-   Download/Clone this project to a location on your system.
-   Open Command Prompt (Windows) or Terminal (MacOS/Linux).
-   Navigate to this folder
    `cd /path/to/this/project/Church_PPT_Utility`
-   Install required packages by running:
    `pip install -r requirements.txt`
-   Once thats installed, you can run the utility by running `python3 ppt_utility.py` and follow the onscreen instructions
-   Once done the updated PPT will be in the `output_ppt/updated_ppt_for_service.pptx` location
-   The bible portions will be in the `output_bible_portions` folder.

## Enhancements for Future

-   Currently there is no usable dataset for our song books so lyrics retreival is not possible. If it is available, adding lyrics can be implemented
-   Currently we are not displaying the bible portions in the PPT because python-pptx does not have an option to insert slides in between the presentation. Hence we are extracting the verses into the `output_bible_portions` folder. Once python-pptx provides that feature, we can think of adding the verses directly in the PPT
