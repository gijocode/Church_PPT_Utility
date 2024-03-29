import datetime
import json
import logging
import os
import platform
from collections import abc
from copy import deepcopy

import six
from lxml import etree
from pptx import Presentation, exc
from pptx.enum.text import MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Pt
from pptx.dml.color import RGBColor

from BibleExtractor import BibleConverter

logger = logging.getLogger("ppt_util")
fh = logging.FileHandler("ppt_log.log")
fh.setLevel(logging.INFO)
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
fh.setFormatter(formatter)
logger.addHandler(fh)
logger.setLevel(logging.INFO)
operating_sys = platform.system()


class PresentationBuilder(object):
    REQ_SLIDES = [
        "first_slide",
        "communion",
        "doxology",
        "offertory",
        "birthday",
        "gospel",
        "epistle",
        "second_lesson",
        "between_lessons",
        "first_lesson",
        "opening_song",
        "template_bible_verse",
        "template_bible_heading",
        "template_song_verse",
    ]

    def __init__(self, service_type) -> None:
        logger.info("Created ppt object")
        logger.info(f"Creating PPT on {datetime.date.today()}")
        template = f"templates/{service_type}_template.pptx"
        self.presentation = Presentation(template)
        self.template_dict = {
            shape.name: slide
            for slide in self.presentation.slides
            for shape in slide.shapes
            if shape.name in self.REQ_SLIDES
        }

        with open("assets/kk_full.json", "r", encoding="utf-8") as kkfile:
            self.kk_dict = json.load(kkfile)
        self.bible_converter = BibleConverter()

    def duplicate_slide(self, template_slide):
        blank_slide_layout = self.presentation.slide_layouts[
            len(self.presentation.slide_layouts) - 1
        ]

        copied_slide = self.presentation.slides.add_slide(blank_slide_layout)

        for shape in template_slide.shapes:
            el = shape.element
            newel = deepcopy(el)
            copied_slide.shapes._spTree.insert_element_before(newel, "p:extLst")

        for _, value in six.iteritems(template_slide.part.rels):
            # Make sure we don't copy a notesSlide relation as that won't exist
            if "notesSlide" not in value.reltype:
                copied_slide.part.rels.add_relationship(
                    value.reltype, value._target, value.rId
                )

        return copied_slide

    @property
    def xml_slides(self):
        return self.presentation.slides._sldIdLst  # pylint: disable=protected-access

    def move_slide(self, old_index, new_index):
        slides = list(self.xml_slides)
        self.xml_slides.remove(slides[old_index])
        self.xml_slides.insert(new_index, slides[old_index])

    # also works for deleting slides
    def delete_slide(self, index):
        slides = list(self.xml_slides)
        self.xml_slides.remove(slides[index])

    def get_kk_lyrics(self, song_no):
        lyrics = self.kk_dict[song_no]["Song"]
        lyrics_slide = lyrics.split("\n\n")
        if lyrics_slide[1].startswith("1"):
            chorus = lyrics_slide[0]
            for i in range(2, 2 * len(lyrics_slide), 2):
                lyrics_slide.insert(i, chorus)
        song_name = self.kk_dict[song_no]["Name"]
        print(song_name)
        title = f"Song No: {self.kk_dict[song_no]['No']}\n{song_name.split('(')[0]}"
        if author := self.kk_dict[song_no].get("Author"):
            title += f"\n\nComposer: {author}"
        lyrics_slide.insert(0, title)
        return lyrics_slide

    def get_bible_portions(self, bib_portion, gui=False):
        if not gui:
            print("\033c", end="")
            portion_type = bib_portion
            print(portion_type.replace("_", " ").upper())
            (
                book,
                chapter,
                starting_verse,
                ending_verse,
            ) = self.bible_converter.get_input_from_user(portion_type)

            book += (
                39
                if portion_type in ["second_lesson", "gospel"]
                else 43 if portion_type == "epistle" else 0
            )
        else:
            book, chapter, starting_verse, ending_verse, portion_type = (
                self.bible_converter.bible_books_eng.index(bib_portion["book"]),
                int(bib_portion["chapter"]) - 1,
                int(bib_portion["starting_verse"]) - 1,
                int(bib_portion["ending_verse"]) - 1,
                bib_portion["type"],
            )

        if portion_type == "all":
            portion_type = ""

        # book, chapter, starting_verse, ending_verse = 1, 1, 2, 5
        bible_portion = [
            "{5}\n\n{4} {1}: {2}-{3}\n{0} {1}: {2}-{3}".format(
                self.bible_converter.bible_books_eng[book],
                chapter + 1,
                starting_verse + 1,
                ending_verse + 1,
                self.bible_converter.bible_books_mal[book],
                portion_type.upper().replace("_", " "),
            )
        ]
        if not portion_type:
            bible_portion[0] = bible_portion[0].strip("\n")
        portion = [
            "{}\n\n{}".format(
                self.bible_converter.extract_bible_portion(
                    self.bible_converter.malbible, book, chapter, i
                ),
                self.bible_converter.extract_bible_portion(
                    self.bible_converter.engbible, book, chapter, i
                ),
            )
            for i in range(starting_verse, ending_verse + 1)
        ]

        logger.info(
            f"Retrieved bible portion for {portion_type} which is {bible_portion}"
        )
        self.put_in_ppt(portion, "template_bible_verse", portion_type, 1)
        self.put_in_ppt(bible_portion, "template_bible_heading", portion_type, 1)

    def update_first_slide(self, topic):
        first_slide = self.template_dict["first_slide"]
        next_sunday = datetime.date.today() + datetime.timedelta(
            days=(6 - datetime.date.today().weekday() + 7) % 7
        )
        for shape in first_slide.shapes:
            if shape.name in ["theme", "todays_date"]:
                tf = shape.text_frame
                tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
                tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
                for para in tf.paragraphs:
                    para.text = (
                        topic
                        if shape.name == "theme"
                        else next_sunday.strftime("%-d %B, %Y")
                    )
                    para.font.name = (
                        "Goudy Bookletter 1911" if shape.name == "theme" else "Arial"
                    )
                    para.font.size = Pt(50)
                    para.font.color.rgb = (
                        RGBColor(0, 13, 200)
                        if shape.name == "theme"
                        else RGBColor(235, 114, 8)
                    )
                    para.font.bold = True
                    para.alignment = PP_ALIGN.CENTER

    def get_song(self, song_type, gui=False):
        if not gui:
            songs = input(f"\nEnter songs for {song_type.replace('_',' ')}: ").split(
                ","
            )[::-1]
            for song in songs:
                if not song.isdigit() and not song.startswith("dox"):
                    song_lyrics = [song]
                else:
                    song_lyrics = self.get_kk_lyrics(song)
                self.put_in_ppt(song_lyrics, "template_song_verse", song_type, 1.5)
        else:
            song_nos, song_type = song_type["no"], song_type["type"]
            songs = song_nos.split(",")[::-1]
            for song_no in songs:
                song_lyrics = self.get_kk_lyrics(song_no)
                self.put_in_ppt(song_lyrics, "template_song_verse", song_type, 1.5)

    def put_in_ppt(
        self,
        content_list: list[str],
        template_name,
        after_slide,
        line_spacing: float = 1,
    ):
        content_slides = []
        for content in content_list[::-1]:
            verse_slide = self.duplicate_slide(self.template_dict[template_name])
            for shape in verse_slide.shapes:
                if shape.name == template_name:
                    tf = shape.text_frame
                    tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
                    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
                    tf.clear()
                    for para in tf.paragraphs:
                        para.text = content
                        if content.count("\n") < 6:
                            para.line_spacing = line_spacing

                        if template_name == "template_bible_heading":
                            para.font.name = "Noto Serif Malayalam"
                            para.font.size = Pt(54)
                            para.font.color.rgb = RGBColor(255, 255, 255)
                        else:
                            para.font.size = Pt(40)
                            para.font.name = "Goudy Bookletter 1911"
                        para.font.bold = True
                        # Hacky workaround to work with malayalam fonts
                        defrpr = para._element.pPr.defRPr
                        ea = etree.SubElement(defrpr, qn("a:cs"))
                        ea.set("typeface", "Noto Serif Malayalam")
                        para.alignment = PP_ALIGN.CENTER

            content_slides.append(verse_slide)
        index_of_slide = (
            self.presentation.slides.index(self.template_dict[after_slide])
            if after_slide
            else 0
        )
        for i in range(len(content_slides)):
            self.move_slide(-1, index_of_slide + 1 + i)

    def save_ppt(self, ppt_name=None):
        if operating_sys == "Darwin":
            next_sunday = (
                datetime.date.today()
                + datetime.timedelta(days=(6 - datetime.date.today().weekday() + 7) % 7)
            ).strftime("%-d %B, %Y")
            # Change the below path to where you want the ppt saved
            ppt_name = (
                (
                    f"/Users/gijomathew/Important/misc/Church/PPTs/2024/{next_sunday}.pptx"
                )
                if not ppt_name
                else ppt_name
            )
            self.presentation.save(ppt_name)
            os.system("open '/Users/gijomathew/Important/misc/Church/PPTs/2024/'")
            os.system(f"open '{ppt_name}'")
        else:
            ppt_name = "holy_communion.pptx" if not ppt_name else ppt_name
            self.presentation.save(ppt_name)
        print(f"PPT Successfully saved to {ppt_name}")
        logger.info(f"PPT Successfully saved to {ppt_name}")


if __name__ == "__main__":
    choice = int(
        input(
            "Enter your choice \n1.Malayalam Service PPT\n2.English Service PPT\n3.Bible Portion PPT\n"
        )
    )

    if choice == 1:
        service_type = "malayalam_service"
    elif choice == 2:
        service_type = "english_service"
    elif choice == 3:
        service_type = "bible_portion"
    else:
        raise ValueError("Invalid choice")
    pb_obj = PresentationBuilder(service_type)

    if service_type == "bible_portion":
        pb_obj.get_bible_portions("all")
        pb_obj.save_ppt("bible_portion.pptx")
        exit(0)
    _ = [
        pb_obj.get_bible_portions(portion)
        for portion in ["first_lesson", "second_lesson", "epistle", "gospel"]
    ]
    print("\033c", end="")
    print(
        "\nNote\n1. If multiple songs, enter comma separated. \n2. For doxology 1 enter dox1 and so on\n3. If song is from a book other than Kriteeya Geethangal, just type the song heading. (Lyrics won't be provided)"
    )
    _ = [
        pb_obj.get_song(song)
        for song in (
            "opening_song",
            "between_lessons",
            "birthday",
            "offertory",
            "communion",
            "doxology",
        )
    ]
    # In Windows sometimes the first slide fails
    try:
        pb_obj.update_first_slide(input("\nEnter the theme for this Sunday: "))
    except Exception as e:
        print(
            "There was some error encountered when updating the first slide. Please make sure you edit it manually!"
        )
        logger.exception(e)

    pb_obj.save_ppt()
