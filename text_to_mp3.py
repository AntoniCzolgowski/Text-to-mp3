import os
import time
from docx import Document
from gtts import gTTS
from moviepy.editor import concatenate_audioclips, AudioFileClip
from pydub import AudioSegment

#audio
def generate_polish_audio(text, filename):
    tts = gTTS(text=text, lang='pl')
    tts.save(filename)
def generate_english_audio(text, filename):
    tts = gTTS(text=text, lang='en')
    tts.save(filename)
#silent
def create_silence(duration, filename):
    silence = AudioSegment.silent(duration=duration * 1000)
    silence.export(filename, format="mp3")

#safe filename
def safe_filename(text):
    text = text.replace('/', '_lub_')
    text = text.replace(' ', '_')
    text = text.replace(',', '_')
    text = text.replace('(', '_').replace(')', '_')
    return text

#docx
doc = Document("") #<- place for the file path

polish_audio_files = []
english_audio_files = []
silence_files = []

#iteration over whole doc
for table in doc.tables:
    for row in table.rows:
        english_word = row.cells[0].text.strip()
        polish_word = row.cells[1].text.strip()

        
        safe_polish_word = safe_filename(polish_word)
        safe_english_word = safe_filename(english_word)

        polish_audio_file = f"polish_{safe_polish_word}.mp3"
        english_audio_file = f"english_{safe_english_word}.mp3"
        silence_file = f"silence_{safe_english_word}.mp3"

        #audio
        generate_polish_audio(polish_word, polish_audio_file)
        time.sleep(2) 
        generate_english_audio(english_word, english_audio_file)
        create_silence(5, silence_file) #5 seconds of silence

        #list
        polish_audio_files.append(polish_audio_file)
        english_audio_files.append(english_audio_file)
        silence_files.append(silence_file)

#combining audio
audio_clips = []
for i in range(len(polish_audio_files)):
    polish_clip = AudioFileClip(polish_audio_files[i])
    english_clip = AudioFileClip(english_audio_files[i])
    silence_clip = AudioFileClip(silence_files[i])

    #polish + silence + english
    combined_clip = concatenate_audioclips([polish_clip, silence_clip, english_clip])
    audio_clips.append(combined_clip)

#combining all the audio clips
final_audio = concatenate_audioclips(audio_clips)

#mp3 file
final_audio.write_audiofile("output.mp3")


