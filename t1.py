import streamlit as st
import re
from pytube import YouTube
from pytube.exceptions import VideoUnavailable, RegexMatchError
from faster_whisper import WhisperModel
import tempfile
import os
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.units import inch
from io import BytesIO

def extract_video_id(url):
    if re.match(r'^[a-zA-Z0-9_-]{11}$', url):
        return url
    patterns = [
        r'youtube\.com/watch\?v=([a-zA-Z0-9_-]{11})',
        r'youtu\.be/([a-zA-Z0-9_-]{11})',
        r'youtube\.com/embed/([a-zA-Z0-9_-]{11})'
    ]
    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)
    return None

def download_audio(video_url):
    try:
        yt = YouTube(video_url)
        audio_stream = yt.streams.filter(only_audio=True).order_by('abr').desc().first()
        if audio_stream is None:
            st.error("No audio stream found.")
            return None, None
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".mp4")
        audio_stream.download(filename=temp_file.name)
        return temp_file.name, yt.title
    except VideoUnavailable:
        st.error("Video unavailable.")
        return None, None
    except RegexMatchError:
        st.error("Invalid YouTube URL or ID.")
        return None, None
    except Exception as e:
        st.error(f"Error downloading audio: {str(e)}")
        return None, None

@st.cache_resource(show_spinner=False)
def load_whisper_model():
    # Use compute_type="int8" for compatibility on CPU or unsupported GPUs
    return WhisperModel("base", device="auto", compute_type="int8")

def transcribe_audio(model, audio_path):
    segments, info = model.transcribe(audio_path)
    text = ""
    for segment in segments:
        text += segment.text.strip() + " "
    return text.strip()

def generate_pdf(text, title):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter,
                            rightMargin=40, leftMargin=40,
                            topMargin=60, bottomMargin=40)
    styles = getSampleStyleSheet()
    story = []

    title_style = styles['Heading1']
    title_style.textColor = "#FF0000"
    story.append(Paragraph(f"Transcript for: {title}", title_style))
    story.append(Spacer(1, 0.25 * inch))

    para_style = ParagraphStyle(
        'normal',
        parent=styles['Normal'],
        fontSize=11,
        leading=15,
        spaceAfter=12,
    )

    sentences = [s.strip() for s in text.split('. ') if s.strip()]
    for sentence in sentences:
        if not sentence.endswith('.'):
            sentence += '.'
        story.append(Paragraph(sentence, para_style))

    doc.build(story)
    buffer.seek(0)
    return buffer

def get_video_streams(video_url):
    try:
        yt = YouTube(video_url)
        streams = yt.streams.filter(progressive=True).order_by('resolution').desc()
        stream_list = []
        for s in streams:
            stream_list.append({
                "itag": s.itag,
                "resolution": s.resolution,
                "mime_type": s.mime_type,
                "filesize_mb": round(s.filesize / (1024*1024), 2) if s.filesize else None
            })
        return yt.title, stream_list
    except Exception as e:
        st.error(f"Error fetching video streams: {str(e)}")
        return None, []

def download_video(video_url, itag):
    try:
        yt = YouTube(video_url)
        stream = yt.streams.get_by_itag(itag)
        if stream is None:
            st.error("Selected stream not found.")
            return None
        buffer = BytesIO()
        stream.stream_to_buffer(buffer)
        buffer.seek(0)
        return buffer
    except Exception as e:
        st.error(f"Error downloading video: {str(e)}")
        return None

st.title("YouTube Video Transcript & Download with Faster Whisper")

video_url_input = st.text_input("Enter YouTube Video URL or Video ID:")

if video_url_input:
    video_id = extract_video_id(video_url_input)
    if video_id:
        video_url = f"https://www.youtube.com/watch?v={video_id}"
        thumbnail_url = f"https://i.ytimg.com/vi/{video_id}/maxresdefault.jpg"
        st.image(thumbnail_url, caption="Video Thumbnail", use_container_width=True)

        model = load_whisper_model()

        with st.spinner("Downloading audio and transcribing with Faster Whisper... This may take a while."):
            audio_path, video_title = download_audio(video_url)
            if audio_path:
                transcript_text = transcribe_audio(model, audio_path)
                os.unlink(audio_path)  # Remove temp audio file

                st.subheader(f"Video Title: {video_title}")
                st.subheader("Transcript:")
                st.write(transcript_text)

                pdf_buffer = generate_pdf(transcript_text, video_title)
                st.download_button(
                    label="Download Transcript as PDF",
                    data=pdf_buffer,
                    file_name=f"{video_title}_transcript.pdf",
                    mime="application/pdf"
                )

                st.subheader("Download Video")
                video_title, streams = get_video_streams(video_url)
                if streams:
                    options = [f"{s['resolution']} - {s['mime_type']} - {s['filesize_mb']} MB" for s in streams]
                    selected = st.selectbox("Select quality:", options)
                    selected_itag = streams[options.index(selected)]["itag"]

                    if st.button("Download Video"):
                        with st.spinner("Downloading video..."):
                            video_buffer = download_video(video_url, selected_itag)
                            if video_buffer:
                                st.download_button(
                                    label="Click here to download video",
                                    data=video_buffer,
                                    file_name=f"{video_title}_{selected}.mp4",
                                    mime="video/mp4"
                                )
                else:
                    st.info("No downloadable video streams found.")
            else:
                st.error("Failed to download audio for transcription.")
    else:
        st.error("Invalid YouTube video URL or ID.")
