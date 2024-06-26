import json, os
from pyfzf.pyfzf import FzfPrompt


class BibleExtractor:

    def __init__(self):
        self.bible_books_mal = [
            "ഉല്പത്തി",
            "പുറപ്പാടു്",
            "ലേവ്യപുസ്തകം",
            "സംഖ്യാപുസ്തകം",
            "ആവർത്തനം",
            "യോശുവ",
            "ന്യായാധിപന്മാർ",
            "രൂത്ത്",
            "1 ശമൂവേൽ",
            "2 ശമൂവേൽ",
            "1 രാജാക്കന്മാർ",
            "2 രാജാക്കന്മാർ",
            "1 ദിനവൃത്താന്തം",
            "2 ദിനവൃത്താന്തം",
            "എസ്രാ",
            "നെഹെമ്യാവു",
            "എസ്ഥേർ",
            "ഇയ്യോബ്",
            "സങ്കീർത്തനങ്ങൾ",
            "സദൃശ്യവാക്യങ്ങൾ",
            "സഭാപ്രസംഗി",
            "ഉത്തമഗീതം",
            "യെശയ്യാ",
            "യിരേമ്യാവു",
            "വിലാപങ്ങൾ",
            "യേഹേസ്കേൽ",
            "ദാനീയേൽ",
            "ഹോശേയ",
            "യോവേൽ",
            "ആമോസ്",
            "ഓബദ്യാവു",
            "യോനാ",
            "മീഖാ",
            "നഹൂം",
            "ഹബക്കൂക്‍",
            "സെഫന്യാവു",
            "ഹഗ്ഗായി",
            "സെഖർയ്യാവു",
            "മലാഖി",
            "മത്തായി",
            "മർക്കൊസ്",
            "ലൂക്കോസ്",
            "യോഹന്നാൻ",
            "പ്രവര്‍ത്തനങ്ങള്‍",
            "റോമർ",
            "1 കൊരിന്ത്യർ",
            "2 കൊരിന്ത്യർ",
            "ഗലാത്യർ",
            "എഫെസ്യർ",
            "ഫിലിപ്പിയർ",
            "കൊലൊസ്സ്യർ",
            "1 തെസ്സലൊനീക്യർ",
            "2 തെസ്സലൊനീക്യർ",
            "1 തിമൊഥെയൊസ്",
            "2 തിമൊഥെയൊസ്",
            "തീത്തൊസ്",
            "ഫിലേമോൻ",
            "എബ്രായർ",
            "യാക്കോബ്",
            "1 പത്രൊസ്",
            "2 പത്രൊസ്",
            "1 യോഹന്നാൻ",
            "2 യോഹന്നാൻ",
            "3 യോഹന്നാൻ",
            "യൂദാ",
            "വെളിപാട",
        ]
        self.bible_books_eng = [
            "Genesis",
            "Exodus",
            "Leviticus",
            "Numbers",
            "Deuteronomy",
            "Joshua",
            "Judges",
            "Ruth",
            "1 Samuel",
            "2 Samuel",
            "1 Kings",
            "2 Kings",
            "1 Chronicles",
            "2 Chronicles",
            "Ezra",
            "Nehemiah",
            "Esther",
            "Job",
            "Psalms",
            "Proverbs",
            "Ecclesiastes",
            "Song of Solomon",
            "Isaiah",
            "Jeremiah",
            "Lamentations",
            "Ezekiel",
            "Daniel",
            "Hosea",
            "Joel",
            "Amos",
            "Obadiah",
            "Jonah",
            "Micah",
            "Nahum",
            "Habakkuk",
            "Zephaniah",
            "Haggai",
            "Zechariah",
            "Malachi",
            "Matthew",
            "Mark",
            "Luke",
            "John",
            "Acts",
            "Romans",
            "1 Corinthians",
            "2 Corinthians",
            "Galatians",
            "Ephesians",
            "Philippians",
            "Colossians",
            "1 Thessalonians",
            "2 Thessalonians",
            "1 Timothy",
            "2 Timothy",
            "Titus",
            "Philemon",
            "Hebrews",
            "James",
            "1 Peter",
            "2 Peter",
            "1 John",
            "2 John",
            "3 John",
            "Jude",
            "Revelation",
        ]
        self.old_testament = self.bible_books_eng[:39]
        self.new_testament = self.bible_books_eng[39:]
        self.gospels = self.new_testament[:4]
        self.epistles = self.new_testament[4:-1]
        self.fzf = FzfPrompt() if not os.system('fzf --version') else None

        self.mapper = {
            "all": self.bible_books_eng,
            "first_lesson": self.old_testament,
            "second_lesson": self.new_testament,
            "epistle": self.epistles,
            "gospel": self.gospels,
        }
        with open("assets/malayalam_bible.json", "r", encoding="utf-8") as bible:
            self.malbible = json.load(bible)

        with open("assets/english_bible.json", "r", encoding="utf-8") as bible:
            self.engbible = json.load(bible)

    def extract_bible_portion(self, bible, book, chapter, verse):
        try:
            return f'{verse+1} {bible["Book"][book]["Chapter"][chapter]["Verse"][verse]["Verse"]}'
        except IndexError:
            return ""

    def get_input_from_user(self, book_type="all"):
        books = self.mapper[book_type]

        if not self.fzf:
            book_printer = []
            formatter_template = "{0:20}|{1:20}|{2:20}|{3:20}|{4:20}"
            for i, book in enumerate(books):
                book_printer.append(f"{i+1}. {book}")
                if len(book_printer) == 5:
                    print(formatter_template.format(*book_printer))
                    book_printer = []
            for _ in range(5 - len(book_printer)):
                book_printer.append("")
            print(formatter_template.format(*book_printer))
            book = int(input("\nEnter index of Book: ")) - 1
        else:
            book_name = self.fzf.prompt(
                books,
                "--layout=reverse-list --height=~40% --border=bold --no-info",
            )[0]
            book = books.index(book_name)
            print("Book:", book_name)

        chapter = int(input("\nEnter the chapter: ")) - 1
        starting_verse = int(input("\nEnter the starting verse: ")) - 1
        ending_verse = int(input("\nEnter the ending verse: ")) - 1

        return book, chapter, starting_verse, ending_verse


if __name__ == "__main__":
    ob1 = BibleExtractor()

    print(ob1.epistles)
    print(ob1.gospels)
    print(ob1.new_testament)
    print(ob1.old_testament)
    print(len(ob1.epistles))
