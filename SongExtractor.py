import json, os
from pyfzf.pyfzf import FzfPrompt
import pandas as pd
import requests


class SongExtractor:

    def __init__(self):
        self.fzf = FzfPrompt() if not os.system("fzf --version") else None
        song_index = "assets/texts.csv"
        self.song_index = []
        chunksize = 10**3
        for chunk in pd.read_csv(song_index, chunksize=chunksize):
            for _, row in chunk.iterrows():
                self.song_index.append(f"{row.firstLine} #{row.textAuthNumber}")

    def extract_song_lyrics(self, song_id):
        try:
            response = requests.get(f"https://hymnary.org/api/fulltext/{song_id}")
            response.raise_for_status()
            data = response.json()

            if not data:
                print("No song data found.")
                return []

            song = data[0]
            lyrics = song.get("text", "")
            title = f"Song: {song.get('title', 'Unknown Title')}"

            # Split the lyrics into slides
            lyrics_slides = lyrics.split("\n\r\n\n")

            chorus = ""
            if len(lyrics_slides) > 1 and lyrics_slides[1].startswith("1"):
                chorus = lyrics_slides.pop(0)
            elif len(lyrics_slides) > 1 and lyrics_slides[1].startswith("Refrain"):
                chorus = lyrics_slides.pop(1)

            # Insert chorus before each verse if found
            if chorus:
                for i in range(0, len(lyrics_slides) * 2, 2):
                    lyrics_slides.insert(i, chorus)

            lyrics_slides.insert(0, title)
            lyrics_slides = [slide.replace("\r\n", "") for slide in lyrics_slides]
            return lyrics_slides

        except requests.RequestException as e:
            print(f"Error fetching song lyrics: {e}")
            return []

    def get_input_from_user(self):
        if self.fzf:
            song = self.fzf.prompt(
                self.song_index,
                "--layout=reverse-list --height=~40% --border=bold --no-info",
            )[0]
            song_name, song_id = song.split("#")
            print(song_name, song_id)
            return song_id
        else:
            print("fzf not installed, sorry!")
            return ""

    def get_song(self):
        song_id = self.get_input_from_user()
        return self.extract_song_lyrics(song_id)


if __name__ == "__main__":
    ob1 = SongExtractor()
    song_id = ob1.get_input_from_user()
    ob1.extract_song_lyrics(song_id=song_id)
