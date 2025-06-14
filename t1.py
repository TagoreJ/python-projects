import streamlit as st
import re
from youtube_transcript_api import YouTubeTranscriptApi, TranscriptsDisabled, NoTranscriptFound, VideoUnavailable
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.units import inch
from io import BytesIO
from pytube import YouTube
from pytube.exceptions import VideoUnavailable as PytubeVideoUnavailable, RegexMatchError

def extract_video_id(url):
    """
    Extract the YouTube video ID from a URL or return the input if it looks like an ID.
    Supports URLs like:
    - https://www.youtube.com/watch?v=VIDEO_ID
    - https://youtu.be/VIDEO_ID
    """
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

def get_transcript(video_id, proxies=None):
    """
    Fetch transcript text for the given YouTube video ID using optional proxies.
    """
    try:
        transcript_list = YouTubeTranscriptApi.get_transcript(video_id, proxies=proxies)
        transcript_text = " ".join([entry['text'] for entry in transcript_list])
        return transcript_text
    except TranscriptsDisabled:
        return "Transcripts are disabled for this video."
    except NoTranscriptFound:
        return "No transcript found for this video."
    except VideoUnavailable:
        return "The video is unavailable."
    except Exception as e:
        return f"An error occurred: {str(e)}"

def generate_pdf(transcript_text, video_id, video_title):
    """
    Generate a PDF from transcript text and return a BytesIO buffer.
    """
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter,
                            rightMargin=40, leftMargin=40,
                            topMargin=60, bottomMargin=40)
    styles = getSampleStyleSheet()
    story = []

    title_style = styles['Heading1']
    title_style.textColor = "#FF0000"  # YouTube red accent
    story.append(Paragraph(f"Transcript for: {video_title}", title_style))
    story.append(Spacer(1, 0.25 * inch))

    para_style = ParagraphStyle(
        'transcript',
        parent=styles['Normal'],
        fontSize=11,
        leading=15,
        spaceAfter=12,
    )

    sentences = [s.strip() for s in transcript_text.split('. ') if s.strip()]
    for sentence in sentences:
        if not sentence.endswith('.'):
            sentence += '.'
        story.append(Paragraph(sentence, para_style))

    doc.build(story)
    buffer.seek(0)
    return buffer

def get_video_streams(video_id):
    """
    Return list of progressive streams (video+audio) with resolution and itag.
    """
    try:
        yt = YouTube(f"https://www.youtube.com/watch?v={video_id}")
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
    except (PytubeVideoUnavailable, RegexMatchError):
        st.error("Video unavailable or invalid video ID for download.")
        return None, []
    except Exception as e:
        st.error(f"Error fetching video streams: {str(e)}")
        return None, []

def download_video(video_id, itag):
    """
    Download video stream by itag and return bytes.
    """
    try:
        yt = YouTube(f"https://www.youtube.com/watch?v={video_id}")
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

st.title("YouTube Video Transcript & Download with Proxy Support")

# Proxy inputs (optional)
st.sidebar.header("Proxy Settings (Optional)")
http_proxy = st.sidebar.text_input("HTTP Proxy (e.g. http://user:pass@host:port)", "")
https_proxy = st.sidebar.text_input("HTTPS Proxy (e.g. https://user:pass@host:port)", "")

proxies = None
if http_proxy or https_proxy:
    proxies = {}
    if http_proxy:
        proxies["http"] = http_proxy
    if https_proxy:
        proxies["https"] = https_proxy

video_url = st.text_input("Enter YouTube Video URL or Video ID:")

if video_url:
    video_id = extract_video_id(video_url)
    if video_id:
        st.write(f"Extracted Video ID: **{video_id}**")

        thumbnail_url = f"https://i.ytimg.com/vi/{video_id}/maxresdefault.jpg"
        st.image(thumbnail_url, caption="Video Thumbnail", use_container_width=True)

        with st.spinner("Fetching video info..."):
            video_title, streams = get_video_streams(video_id)

        if video_title:
            st.subheader(f"Video Title: {video_title}")
        else:
            st.error("Could not fetch video title.")

        with st.spinner("Fetching transcript..."):
            transcript = get_transcript(video_id, proxies=proxies)

        st.subheader("Transcript:")
        st.write(transcript)

        if not transcript.startswith(("Transcripts are disabled", "No transcript found", "The video is unavailable", "An error occurred")):
            pdf_buffer = generate_pdf(transcript, video_id, video_title if video_title else video_id)
            st.download_button(
                label="Download Transcript as PDF",
                data=pdf_buffer,
                file_name=f"{video_id}_transcript.pdf",
                mime="application/pdf"
            )

        if streams:
            st.subheader("Download Video")
            options = [f"{s['resolution']} - {s['mime_type']} - {s['filesize_mb']} MB" for s in streams]
            selected = st.selectbox("Select quality:", options)

            selected_itag = streams[options.index(selected)]["itag"]

            if st.button("Download Video"):
                with st.spinner("Downloading video..."):
                    video_buffer = download_video(video_id, selected_itag)
                    if video_buffer:
                        st.download_button(
                            label="Click here to download video",
                            data=video_buffer,
                            file_name=f"{video_title if video_title else video_id}_{selected}.mp4",
                            mime="video/mp4"
                        )
    else:
        st.error("Invalid YouTube video URL or ID. Please enter a valid one.")
