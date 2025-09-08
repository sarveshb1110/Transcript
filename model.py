import subprocess
import os
import re
from faster_whisper import WhisperModel
from docx import Document


# --------- CONFIGURATION ---------
video_path = r"C:\Users\SARVESH BIDWE\Quantian\Speech2Text\EVA0313-0001 1 (1).mkv"  # Update your video path here
audio_path = "temp_audio.wav"


base_filename = os.path.splitext(os.path.basename(video_path))[0]
prefix = base_filename[:8].replace(" ", "_")
txt_filename = f"{prefix}_transcription.txt"
docx_filename = f"{prefix}_transcription.docx"


# --------- STEP 1: Extract Audio from Video ---------
print("Extracting audio from video...")
ffmpeg_command = [
    "ffmpeg", "-y", "-i", video_path,
    "-vn", "-acodec", "pcm_s16le", "-ar", "16000", "-ac", "1", audio_path
]
subprocess.run(ffmpeg_command, check=True)


# --------- STEP 2: Transcribe Audio ---------
print("Loading Whisper model...")
model = WhisperModel("base", device="cpu")
print("Transcribing audio...")
segments, info = model.transcribe(audio_path, language=None)


raw_transcription = ""
for segment in segments:
    raw_transcription += segment.text.strip() + " "


os.remove(audio_path)
print("Temporary audio file deleted.")


# --------- STEP 3: Enhanced Text Cleaning ---------
def clean_text(text):
    text = text.lower()

    # Replace spoken punctuation words with punctuation marks
    text = re.sub(r"\b(full\s*stop|stop|full-stop)\b", ".", text)
    text = re.sub(r"\b(comma|coma)\b", ",", text)

    # Replace spoken paragraph words with paragraph breaks
    text = re.sub(r"\b(new paragraph|para|paragraph)\b", "\n\n", text)

    # Currency formatting: convert currency words + number → symbol + number
    currency_map = {
        "dollars": r"$",
        "dollar": r"$",
        "grand": r"$",
        "pounds": r"£",
        "pound": r"£",
        "euros": r"€",
        "euro": r"€",
        "rupees": r"₹",
        "rupee": r"₹",
    }

    # Regex to find patterns like "dollars 500", "pounds 200.50"
    def replace_currency(match):
        currency_word = match.group(1)
        amount = match.group(2)
        symbol = currency_map.get(currency_word, "")
        return f"{symbol}{amount}"

    text = re.sub(
        r"\b(dollars?|pounds?|euros?|rupees?)\s+(\d+(\.\d+)?)\b",
        replace_currency,
        text,
    )

    # Fix consecutive punctuation marks
    text = re.sub(r"(\.)(\s*\.)+", ".", text)  # multiple dots → one dot
    text = re.sub(r",\s*\.", ".", text)
    text = re.sub(r"\.\s*,", ",", text)
    text = re.sub(r",\s*,", ",", text)

    # Remove extra spaces before/after punctuation
    text = re.sub(r"\s+([.,])", r"\1", text)
    text = re.sub(r"([.,])\s+", r"\1 ", text)

    # Capitalize first letter of paragraphs and sentences
    paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
    cleaned_paragraphs = []
    for para in paragraphs:
        para = para[0].upper() + para[1:] if para else ""
        para = re.sub(
            r'([.!?]\s+)([a-z])', lambda m: m.group(1) + m.group(2).upper(), para
        )
        cleaned_paragraphs.append(para)

    return "\n\n".join(cleaned_paragraphs).strip()


# Apply the cleaning function and get the cleaned transcription text
cleaned_text = clean_text(raw_transcription)


# --------- STEP 4: Save cleaned text to .txt file ---------
with open(txt_filename, "w", encoding="utf-8") as f:
    f.write(cleaned_text)
print(f"Cleaned transcription saved to text file: {txt_filename}")


# --------- STEP 5: Save cleaned text to Word (.docx) file ---------
doc = Document()
for para in cleaned_text.split("\n\n"):
    doc.add_paragraph(para)
doc.save(docx_filename)
print(f"Cleaned transcription saved to Word file: {docx_filename}")
