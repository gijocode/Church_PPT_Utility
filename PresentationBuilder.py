import datetime
import json
import os
from collections import abc
from copy import deepcopy

import six
from lxml import etree
from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn

from BibleExtractor import BibleConverter


class PresentationBuilder(object):
    REQ_SLIDES = [
        "first_slide",
        "communion",
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

    def __init__(self) -> None:
        self.presentation = Presentation("templates/malayalam_service_template.pptx")
        self.template_dict = {}
        for slide in self.presentation.slides:
            for shape in slide.shapes:
                if shape.name == "maltext":
                    tf = shape.text_frame
                    for p in tf.paragraphs:
                        with open("xmlslide.xml", "w") as xmlfile:
                            xmlfile.write(p._element.xml)
                if shape.name in self.REQ_SLIDES:
                    self.template_dict[shape.name] = slide
        with open("assets/kk_full.json", "r") as kkfile:
            self.kk_dict = json.load(kkfile)
        self.bible_converter = BibleConverter()

    def duplicate_slide(self, template_slide):
        try:
            blank_slide_layout = self.presentation.slide_layouts[
                len(self.presentation.slide_layouts) - 1
            ]
        except:
            blank_slide_layout = self.presentation.slide_layouts[
                len(self.presentation.slide_layouts)
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

    def get_bible_portions(self, portion_type):

        os.system("clear")
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
            else 43
            if portion_type == "epistle"
            else 0
        )
        # book, chapter, starting_verse, ending_verse = 1, 1, 2, 5
        bible_portion = [
            "{5}\n\n{0} {1}: {2}-{3}\n\n{4} {1}: {2}-{3}".format(
                self.bible_converter.bible_books_eng[book],
                chapter + 1,
                starting_verse + 1,
                ending_verse + 1,
                self.bible_converter.bible_books_mal[book],
                portion_type.upper().replace("_", " "),
            )
        ]
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

        self.put_in_ppt(portion, "template_bible_verse", portion_type)
        self.put_in_ppt(bible_portion, "template_bible_heading", portion_type)

    def update_first_slide(self, topic):
        first_slide = self.template_dict["first_slide"]
        for shape in first_slide.shapes:
            if shape.name in ["theme", "todays_date"]:
                tf = shape.text_frame
                tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
                tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
                for para in tf.paragraphs:
                    para.text = (
                        topic
                        if shape.name == "theme"
                        else datetime.datetime.now().strftime("%-d %B, %Y")
                    )
                    para.font.name = "Goudy Bookletter 1911"
                    para.alignment = PP_ALIGN.CENTER
                tf.fit_text(font_family="Noto Serif Malayalam", max_size=40)

    def get_song(self, song_type: str):

        no_of_songs = (
            int(input("\nNo of communion songs: ")) if song_type == "communion" else 1
        )
        for _ in range(no_of_songs):
            song_no = input(
                "\nEnter song number for {}:".format(
                    song_type.replace("_", " ").capitalize()
                )
            )
            song_lyrics = self.get_kk_lyrics(song_no)
            self.put_in_ppt(song_lyrics, "template_song_verse", song_type)

    def put_in_ppt(self, content_list, template_name, after_slide):
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
                        para.font.name = "Noto Serif Malayalam"
                        defrpr = para._element.pPr.defRPr
                        ea = etree.SubElement(defrpr, qn("a:cs"))
                        ea.set("typeface", "Noto Serif Malayalam")
                        para.alignment = PP_ALIGN.CENTER
                    try:
                        tf.fit_text(font_family="Noto Serif Malayalam", max_size=40)
                    except:
                        # For some naughty lyrics, that don't fit in the slide
                        tf.fit_text(font_family="Noto Serif Malayalam", max_size=35)

            content_slides.append(verse_slide)
        index_of_slide = self.presentation.slides.index(self.template_dict[after_slide])
        for i in range(len(content_slides)):
            self.move_slide(-1, index_of_slide + 1 + i)

    def save_ppt(self):
        os.system("rm -rf Updated_PPT.pptx")
        self.presentation.save("Updated_PPT.pptx")
        print("PPT Successfullt saved to Updated_PPT.pptx")
        os.system("open Updated_PPT.pptx")


obj1 = PresentationBuilder()
portions = ["first_lesson", "second_lesson", "epistle", "gospel"]
for p in portions:
    obj1.get_bible_portions(p)
os.system("clear")
songs = ["opening_song", "between_lessons", "birthday", "offertory", "communion"]
for song in songs:
    obj1.get_song(song)
obj1.update_first_slide(input("\nEnter the theme for this Sunday: "))

obj1.save_ppt()