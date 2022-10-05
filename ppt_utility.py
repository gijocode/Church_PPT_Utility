import collections
import collections.abc
import os

from pptx import Presentation
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT

from bible_extractor import BibleExtractor

bible_ext = BibleExtractor()
prs = Presentation("templates/PPT English Service.pptx")
SLIDE_NAMES = [
    "Opening Hymn",
    "First Lesson",
    "Second Lesson",
    "Song between Lessons",
    "Epistle",
    "Gospel",
    "Birthday,Wedding",
    "Offertory",
    "Communion Songs",
    "Doxology",
]
slide_map = {slide.name: slide for slide in prs.slides}

# Getting First lesson notation and verses
notation, verses = bible_ext.get_first_lesson()
slide_map["First Lesson"].shapes[1].text = notation
slide_map["First Lesson"].shapes[1].text_frame.fit_text(
    font_family="Arial", max_size=100
)
slide_map["First Lesson"].shapes[1].text_frame.paragraphs[
    0
].alignment = PP_PARAGRAPH_ALIGNMENT.CENTER  # type: ignore
with open("output_bible_portions/first_lesson.txt", "w") as first_lesson_file:
    first_lesson_file.write(verses)


# Getting Second lesson notation and verses
notation, verses = bible_ext.get_second_lesson()
slide_map["Second Lesson"].shapes[1].text = notation
slide_map["Second Lesson"].shapes[1].text_frame.fit_text(
    font_family="Arial", max_size=100
)
slide_map["Second Lesson"].shapes[1].text_frame.paragraphs[
    0
].alignment = PP_PARAGRAPH_ALIGNMENT.CENTER  # type: ignore
with open("output_bible_portions/second_lesson.txt", "w") as second_lesson_file:
    second_lesson_file.write(verses)

# Getting Epistle notation and verses
notation, verses = bible_ext.get_epistle()
slide_map["Epistle"].shapes[1].text = notation
slide_map["Epistle"].shapes[1].text_frame.fit_text(font_family="Arial", max_size=100)
slide_map["Epistle"].shapes[1].text_frame.paragraphs[
    0
].alignment = PP_PARAGRAPH_ALIGNMENT.CENTER  # type: ignore
with open("output_bible_portions/epistle.txt", "w") as epistle_file:
    epistle_file.write(verses)

# Getting Gospel notation and verses
notation, verses = bible_ext.get_gospel()
slide_map["Gospel"].shapes[1].text = notation
slide_map["Gospel"].shapes[1].text_frame.fit_text(font_family="Arial", max_size=100)
slide_map["Gospel"].shapes[1].text_frame.paragraphs[
    0
].alignment = PP_PARAGRAPH_ALIGNMENT.CENTER  # type: ignore
with open("output_bible_portions/gospel.txt", "w") as gospel_file:
    gospel_file.write(verses)

os.system("clear")  # If using windows, use 'cls'
# Saving final PPT with updated portions and songs
prs.save("output_ppt/test.pptx")
print("Saved the PPT!")
