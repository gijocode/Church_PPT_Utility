import PySimpleGUI as sg

from BibleExtractor import BibleConverter
from PresentationBuilder import PresentationBuilder


def get_inp_from_gui():

    bc = BibleConverter()
    theme_layout = [
        sg.Text(
            "Theme",
            justification="center",
            font="AvenirNext-Regular 26",
        ),
        sg.Input("", key="theme_inp", size=40),
    ]
    bible_layout = [
        sg.Frame(
            "Bible Portions",
            k="biblep",
            font="AvenirNext-Regular 28",
            layout=[
                [sg.Text("")],
                [
                    sg.Text(
                        "First Lesson",
                        font="AvenirNext-Regular 20",
                    ),
                ],
                [
                    sg.Text(
                        "Book",
                    ),
                    sg.Combo(
                        bc.mapper.get("first_lesson"),
                        size=(None, 30),
                        key="bib_fl_book_inp",
                    ),
                    sg.Text(
                        "Chapter",
                    ),
                    sg.Input("", size=3, key="bib_fl_ch_inp"),
                    sg.Text(
                        "Starting Verse",
                    ),
                    sg.Input("", size=3, key="bib_fl_sv_inp"),
                    sg.Text("Ending Verse"),
                    sg.Input("", size=3, key="bib_fl_ev_inp"),
                ],
                [
                    sg.Text(
                        "Second Lesson",
                        font="AvenirNext-Regular 20",
                    ),
                ],
                [
                    sg.Text(
                        "Book",
                    ),
                    sg.Combo(
                        bc.mapper.get("second_lesson"),
                        key="bib_sl_book_inp",
                        size=(None, 30),
                    ),
                    sg.Text("Chapter"),
                    sg.Input("", size=3, key="bib_sl_ch_inp"),
                    sg.Text(
                        "Starting Verse",
                    ),
                    sg.Input("", size=3, key="bib_sl_sv_inp"),
                    sg.Text(
                        "Ending Verse",
                    ),
                    sg.Input("", size=3, key="bib_sl_ev_inp"),
                ],
                [
                    sg.Text(
                        "Epistle",
                        font="AvenirNext-Regular 20",
                    ),
                ],
                [
                    sg.Text(
                        "Book",
                    ),
                    sg.Combo(
                        bc.mapper.get("epistle"),
                        size=(None, 30),
                        key="bib_epi_book_inp",
                    ),
                    sg.Text(
                        "Chapter",
                    ),
                    sg.Input("", size=3, key="bib_epi_ch_inp"),
                    sg.Text(
                        "Starting Verse",
                    ),
                    sg.Input("", size=3, key="bib_epi_sv_inp"),
                    sg.Text(
                        "Ending Verse",
                    ),
                    sg.Input("", size=3, key="bib_epi_ev_inp"),
                ],
                [
                    sg.Text(
                        "Gospel",
                        font="AvenirNext-Regular 20",
                    ),
                ],
                [
                    sg.Text(
                        "Book",
                    ),
                    sg.Combo(
                        bc.mapper.get("gospel"),
                        enable_events=True,
                        key="bib_gos_book_inp",
                        size=(None, 30),
                    ),
                    sg.Text(
                        "Chapter",
                    ),
                    sg.Input("", size=3, key="bib_gos_ch_inp"),
                    sg.Text(
                        "Starting Verse",
                    ),
                    sg.Input("", size=3, key="bib_gos_sv_inp"),
                    sg.Text(
                        "Ending Verse",
                    ),
                    sg.Input("", size=3, key="bib_gos_ev_inp"),
                ],
            ],
        )
    ]
    song_layout = [
        sg.Frame(
            "Songs",
            font="AvenirNext-Regular 28",
            tooltip="Use comma for multiple songs eg: 34,167,87",
            expand_x=True,
            layout=[
                [sg.Text("")],
                [
                    sg.Text(
                        "Opening Song",
                        font="AvenirNext-Regular 20",
                    ),
                    sg.Push(),
                    sg.Input("", size=6, key="song_opening"),
                ],
                [
                    sg.Text(
                        "Between Lessons",
                        font="AvenirNext-Regular 20",
                    ),
                    sg.Push(),
                    sg.Input("", size=6, key="song_lessons"),
                ],
                [
                    sg.Text(
                        "Birthday",
                        font="AvenirNext-Regular 20",
                    ),
                    sg.Push(),
                    sg.Input("", size=6, key="song_bday"),
                ],
                [
                    sg.Text(
                        "Offertory",
                        font="AvenirNext-Regular 20",
                    ),
                    sg.Push(),
                    sg.Input("", size=6, key="song_offer"),
                ],
                [
                    sg.Text(
                        "Communion",
                        font="AvenirNext-Regular 20",
                    ),
                    sg.Push(),
                    sg.Input("", size=12, key="song_comm"),
                ],
                [
                    sg.Text(
                        "Doxology",
                        font="AvenirNext-Regular 20",
                    ),
                    sg.Push(),
                    sg.Input("", size=6, key="song_dox"),
                ],
            ],
        )
    ]

    layout = [
        [
            sg.Push(),
            sg.Text(
                "Emmanuel Mar Thoma Church PPT Builder",
                font=("", 32),
            ),
            sg.Push(),
        ],
        [sg.HorizontalSeparator()],
        [sg.Text("")],
        theme_layout,
        [sg.Text("")],
        [sg.Column([bible_layout]), sg.Column([song_layout])],
        [
            sg.Push(),
            sg.Button("Exit"),
            sg.Button("Create PPT"),
            sg.Push(),
        ],
    ]

    window = sg.Window("EMTC PPT Builder", layout)

    while True:
        event, values = window.read()

        if event == "Create PPT":
            print(values)
            break
        elif event in ("Exit", None):
            break

    window.close()
    return values


if __name__ == "__main__":
    obj1 = PresentationBuilder()
    inputs = get_inp_from_gui()

    bible_portions = [
        {
            "type": "first_lesson",
            "book": inputs["bib_fl_book_inp"],
            "chapter": inputs["bib_fl_ch_inp"],
            "starting_verse": inputs["bib_fl_sv_inp"],
            "ending_verse": inputs["bib_fl_ev_inp"],
        },
        {
            "type": "second_lesson",
            "book": inputs["bib_sl_book_inp"],
            "chapter": inputs["bib_sl_ch_inp"],
            "starting_verse": inputs["bib_sl_sv_inp"],
            "ending_verse": inputs["bib_sl_ev_inp"],
        },
        {
            "type": "epistle",
            "book": inputs["bib_epi_book_inp"],
            "chapter": inputs["bib_epi_ch_inp"],
            "starting_verse": inputs["bib_epi_sv_inp"],
            "ending_verse": inputs["bib_epi_ev_inp"],
        },
        {
            "type": "gospel",
            "book": inputs["bib_gos_book_inp"],
            "chapter": inputs["bib_gos_ch_inp"],
            "starting_verse": inputs["bib_gos_sv_inp"],
            "ending_verse": inputs["bib_gos_ev_inp"],
        },
    ]

    songs = [
        {"type": "opening_song", "no": inputs["song_opening"]},
        {
            "type": "between_lessons",
            "no": inputs["song_lessons"],
        },
        {
            "type": "birthday",
            "no": inputs["song_bday"],
        },
        {
            "type": "offertory",
            "no": inputs["song_offer"],
        },
        {
            "type": "communion",
            "no": inputs["song_comm"],
        },
        {
            "type": "doxology",
            "no": inputs["song_dox"],
        },
    ]

    for portion in bible_portions:
        obj1.get_bible_portions(portion, True)

    for song in songs:
        obj1.get_song(song, True)

    obj1.update_first_slide(inputs["theme_inp"])
    obj1.save_ppt()
