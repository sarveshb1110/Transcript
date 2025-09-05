import subprocess
import os
import re
from pyannote.audio import Pipeline
from faster_whisper import WhisperModel

# --------- CONFIGURATION ---------
video_path = r"C:\Users\SARVESH BIDWE\Quantian\Speech2Text\2508-922 - Bailey & Scott - BAI03080001 - Attendan.wav"  # <<<< Update this path!
audio_path = "temp_audio.wav"

# Automatically build output filename from video
base_filename = os.path.splitext(os.path.basename(video_path))[0]
prefix = base_filename[:8].replace(" ", "_")
final_text_file = f"{prefix}_transcription_final.txt"

# --------- STEP 1: Extract Audio from Video ---------
print("Extracting audio from video...")
ffmpeg_command = [
    "ffmpeg", "-y", "-i", video_path,
    "-vn", "-acodec", "pcm_s16le", "-ar", "16000", "-ac", "1", audio_path
]
subprocess.run(ffmpeg_command, check=True)

# Step 2: Diarization
pipeline = Pipeline.from_pretrained("pyannote/speaker-diarization", use_auth_token="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")
diarization_result = pipeline(audio_path)

# --------- STEP 2: Transcribe Audio to Text ---------
print("Loading Whisper model...")
model = WhisperModel("base", device="cpu")
print("Transcribing audio...")
segments, info = model.transcribe(audio_path)

# --------- STEP 3: Raw Transcription ---------
def get_speaker_label(segment_start, segment_end, diarization):
    for turn, _, speaker in diarization.itertracks(yield_label=True):
        if segment_start >= turn.start and segment_end <= turn.end:
            return speaker
    return "Unknown"

raw_text = ""
for segment in segments:
    speaker = get_speaker_label(segment.start, segment.end, diarization_result)
    raw_text += f"{speaker}: {segment.text.strip()} "

os.remove(audio_path)
print("Temporary audio file deleted.")

# --------- STEP 4: Punctuation Cleanup & Formatting ---------
def clean_text(text):
    # Lowercase for uniform processing
    text = text.lower()

    # Replace punctuation words with symbols
    text = re.sub(r"\b(full\s*stop|stop|full-stop)\b", ".", text)
    text = re.sub(r"\b(comma|coma)\b", ",", text)

    # Split at paragraph spoken markers
    paragraphs = re.split(r"\b(new paragraph|para|paragraph)\b", text, flags=re.IGNORECASE)

    cleaned_paragraphs = []
    for para in paragraphs:
        # Remove any leftover spoken paragraph words and excess whitespace
        para = re.sub(r"\b(new paragraph|para|paragraph)\b", "", para, flags=re.IGNORECASE).strip()

        if not para:
            continue

        # 1. Collapse multiple commas and spaces into one comma
        para = re.sub(r"\s*,\s*,+", ",", para)
        para = re.sub(r",\s*,", ",", para)
        # 2. Collapse multiple periods and spaces into one period
        para = re.sub(r"\s*\.\s*\.+", ".", para)
        para = re.sub(r"\.\s*\.", ".", para)
        # 3. Remove space between comma and period, and replace with single period
        para = re.sub(r",\s*\.", ".", para)
        # 4. Remove space between period and comma, keep as comma
        para = re.sub(r"\.\s*,", ",", para)
        # 5. Remove repeated mix
        para = re.sub(r"\s*([.,])\s*\1+", r"\1", para)

        # 6. Remove starting punctuation (comma or period)
        para = re.sub(r"^[.,\s]+", "", para)
        # 7. Remove ending repeated punctuation, retain only one terminal mark
        para = re.sub(r"[.,\s]+$", "", para)

        # 8. Capitalize the first letter, and after end punctuation + space
        para = para.capitalize()
        para = re.sub(r'([.!?]\s+)(\w)', lambda m: m.group(1) + m.group(2).upper(), para)

        # 9. Ensure paragraph ends with a full stop
        if not para.endswith(('.', '!', '?')):
            para += '.'

        cleaned_paragraphs.append(para)

    # Use double newlines for paragraph separation
    return "\n\n".join(cleaned_paragraphs)
 # Double newlines for true paragraph breaks


cleaned_text = clean_text(raw_text)

with open(final_text_file, "w", encoding="utf-8") as f:
    f.write(cleaned_text)

print(f"Transcription complete! Saved to {final_text_file}")


from docx import Document

# Use the same prefix as your text output for naming
docx_filename = f"{prefix}_transcription_final.docx"

# Read your cleaned text file and split into paragraphs
with open(final_text_file, 'r', encoding='utf-8') as f:
    content = f.read()
    paragraphs = [p.strip() for p in content.split('\n\n') if p.strip()]

# Create and save .docx
doc = Document()
for para in paragraphs:
    doc.add_paragraph(para)
doc.save(docx_filename)

print(f"Word document created: {docx_filename}")


