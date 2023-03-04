import json

kk = {}
with open("assets/kk_full.json", "r") as kkfile:
    kk = json.load(kkfile)


for song_no, song in kk.items():
    lyrics: list = song["Song"].split("\n\n")
    print("song ", song_no)
    if len(lyrics) == 1:
        ll = lyrics[0].split("\n")
        x = len(ll)
        n = 1
        for y in range(4, len(ll) // 2):
            if len(ll) % y == 0:
                n = y
                break
        lyrics = ["\n".join(ll[i : i + n]) for i in range(0, len(ll), n)]
        kk[song_no]["Song"] = "\n\n".join(lyrics)
    if lyrics[0].count("\n") > lyrics[1].count("\n") and not lyrics[1].startswith("1"):
        print(song_no)
        x = lyrics[1].count("\n") + 1
        chorus_lyrics = lyrics[0].split("\n")[x:]
        chorus = "\n".join(chorus_lyrics)
        lyrics[0] = "\n".join(lyrics[0].split("\n")[:x])
        lyrics.insert(0, chorus)
        kk[song_no]["Song"] = "\n\n".join(lyrics)
with open("updated_kk.json", "w") as newkk:
    json.dump(kk, newkk, ensure_ascii=False)

print("NEW FILE")
with open("updated_kk.json", "r") as kkfile:
    kk = json.load(kkfile)


for song_no, song in kk.items():
    lyrics: list = song["Song"].split("\n\n")
    if lyrics[0].count("\n") > lyrics[1].count("\n"):
        print(song_no)
