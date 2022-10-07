"""
    @author Gijo Varughese Mathew; github.com/gijocode
    @year 2022
"""

import json
import os
import re


def clear_console(func):
    def wrapper(self):
        os.system(self.clear_str)
        return func(self)

    return wrapper


class BibleExtractor:
    """
    Utility for extracting verses from the bible.
    """

    def __init__(self, clear_str) -> None:
        """
        King James Version Used.
        Assigning the books for old,new testament, gospels and epistles
        """
        self.clear_str = clear_str
        with open("assets/en_kjv.json", encoding="utf-8-sig") as bible_file:
            self.bible = json.load(
                bible_file,
            )

        self.old_testament = self.bible[:39]  # First 39 books are old testament
        self.new_testament = self.bible[39:]  # Remaining are new testament
        self.gospels = self.new_testament[
            :4
        ]  # First 4 books of new testament are gospels
        self.epistles = self.new_testament[
            5:26
        ]  # Books in the new testament between 5 and 26 are epistles

    def print_books(self, book_list):
        """Gives a CLI visual of the books in the list with indexing

        Args:
            book_list (_type_): old,new,gospel,epistle
        """
        book_names = [x["name"] for x in book_list]
        for i, book in enumerate(book_names):
            if not i % 4:
                print("")
            print(f"{i+1:<2}. {book:<25}", end="")

    def get_misc(self):
        """Miscellaneous Function."""
        pass

    @clear_console
    def get_first_lesson(self):
        print("Select First Lesson")
        book, chapter, verses = self.select_verses(self.old_testament)
        return self.get_verses(book, chapter, verses)

    @clear_console
    def get_second_lesson(self):
        print("Select Second Lesson")
        book, chapter, verses = self.select_verses(self.new_testament)
        return self.get_verses(book, chapter, verses)

    @clear_console
    def get_epistle(self):
        print("Select Epistle")
        book, chapter, verses = self.select_verses(self.epistles)
        return self.get_verses(book, chapter, verses)

    @clear_console
    def get_gospel(self):
        print("Select Gospel")
        book, chapter, verses = self.select_verses(self.gospels)
        return self.get_verses(book, chapter, verses)

    def get_verses(self, book, chapter, verses):
        verse_notation = f"{book} {chapter}:{verses[0]}-{verses[1]}"
        book_dict = next(b for b in self.bible if b["name"] == book)
        pattern = " \{.*?\}"  # Some verses have footnotes. Cleaning them with regex.
        verses_list = [
            f"{verse_no+verses[0]}. {re.sub(pattern,'',verse)}"
            for verse_no, verse in enumerate(
                book_dict["chapters"][chapter - 1][verses[0] - 1 : verses[1]]
            )
        ]  # Guilty of using unreadable comprehension ;)
        verses_text = "\n".join(verses_list)
        # print(verses_text)
        return verse_notation, verses_text

    def select_verses(self, book_list):
        while True:
            self.print_books(book_list)
            book = book_list[int(input("\nEnter yout choice ")) - 1]["name"]
            chapter = int(input(f"Enter the chapter number of {book}: "))
            verses = (
                int(input("Enter starting verse number: ")),
                int(input("Enter ending verse number: ")),
            )

            print(f"You have selected {book} Chapter {chapter}:{verses[0]}-{verses[1]}")
            if input("Proceed? y/n ") == "y":
                break
        return book, chapter, verses


if __name__ == "__main__":
    bible_extract = BibleExtractor("clear")
    bible_extract.get_gospel()
