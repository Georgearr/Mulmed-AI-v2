import os
import time
import difflib
import win32com.client
import speech_recognition as sr
from google.generativeai import configure, GenerativeModel
from dotenv import load_dotenv

#ini apinyaa!!
load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")
configure(api_key=GOOGLE_API_KEY)
model_gemini = GenerativeModel("gemini-pro")

#ini presentasinya!!
ppt = win32com.client.Dispatch("PowerPoint.Application")
pres = ppt.Presentations(1)
slides = pres.Slides

recognizer = sr.Recognizer()
mic = sr.Microphone()

#fungsi untuk ambil text dari slide tertentu
def get_slide_text(index):
    if 1 <= index <= slides.Count:
        text = ""
        for shape in slides(index).Shapes:
            if shape.HasTextFrame:
                if shape.TextFrame.HasText:
                    text += shape.TextFrame.TextRange.Text + " "
        return text.strip()
    return ""

#fungsi untuk minta keputusan dari Gemini
#kt minta AI bandingin ucapan terakhir dengan isi slide

def should_advance_slide(transcript, slide_text):
    prompt = f"Apakah kalimat terakhir dari transkrip berikut cocok atau relevan dengan akhir konten slide ini?\n\nTranskrip: {transcript}\n\nKonten Slide: {slide_text}\n\nJawab dengan 'ya' jika waktunya ganti slide, atau 'tidak' jika belum."
    try:
        response = model_gemini.generate_content(prompt)
        reply = response.text.lower()
        return 'ya' in reply
    except Exception as e:
        print(f"[!] Gemini error: {e}")
        return False

#fungsi utama
#looping terus sambil dengar ucapan dan minta keputusan

def run_slide_changer():
    current_slide = pres.SlideShowWindow.View.CurrentShowPosition
    print("[INFO] Memulai pengenalan suara...")

    while current_slide <= slides.Count:
        with mic as source:
            recognizer.adjust_for_ambient_noise(source)
            print(f"[INFO] Dengarkan untuk slide {current_slide}...")
            audio = recognizer.listen(source)

        try:
            result = recognizer.recognize_google(audio, language="id-ID")
            print(f"[TRANSKRIP] {result}")

            slide_text = get_slide_text(current_slide)

            #kt bandingin sama slide sebelum selanjutnya
            if should_advance_slide(result, slide_text):
                print("[AI] Saatnya pindah slide!")
                pres.SlideShowWindow.View.Next()
                time.sleep(0.75)  #tunggu transisinya
                current_slide += 1
            else:
                print("[AI] Belum waktunya pindah.")

        except sr.UnknownValueError:
            print("[!] Tidak bisa mengenali ucapan.")
        except sr.RequestError as e:
            print(f"[!] Error pada Google Speech Recognition service: {e}")

if __name__ == '__main__':
    run_slide_changer()