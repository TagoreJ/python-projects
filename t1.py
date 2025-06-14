import streamlit as st
from pytube import YouTube
from pytube.exceptions import VideoUnavailable, RegexMatchError

def extract_video_id(url):
    import re
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

st.title("YouTube Video & Audio Downloader (pytube)")

video_url = st.text_input("Enter YouTube Video URL or ID:")

if video_url:
    video_id = extract_video_id(video_url)
    if not video_id:
        st.error("Invalid YouTube URL or Video ID.")
    else:
        full_url = f"https://www.youtube.com/watch?v={video_id}"
        try:
            yt = YouTube(full_url)
        except VideoUnavailable:
            st.error("Video unavailable.")
            st.stop()
        except RegexMatchError:
            st.error("Invalid YouTube URL or Video ID.")
            st.stop()
        except Exception as e:
            st.error(f"Error initializing YouTube object: {str(e)}")
            st.stop()

        st.subheader(f"Title: {yt.title}")
        st.image(yt.thumbnail_url, use_container_width=True)

        # Choose download type
        download_type = st.radio("Select download type:", ("Video", "Audio"))

        if download_type == "Video":
            streams = yt.streams.filter(progressive=True).order_by('resolution').desc()
            options = [stream.resolution for stream in streams]
            selected_res = st.selectbox("Select resolution:", options)

            stream = streams.filter(res=selected_res).first()

            if st.button("Download Video"):
                if stream:
                    with st.spinner("Downloading video..."):
                        file_path = stream.download()
                    st.success(f"Downloaded video: {file_path}")
                    with open(file_path, "rb") as f:
                        st.download_button(
                            label="Click here to download video file",
                            data=f,
                            file_name=f"{yt.title}_{selected_res}.mp4",
                            mime="video/mp4"
                        )
                else:
                    st.error("Selected resolution not available.")

        else:  # Audio download
            streams = yt.streams.filter(only_audio=True).order_by('abr').desc()
            options = [stream.abr for stream in streams]
            selected_abr = st.selectbox("Select audio bitrate:", options)

            stream = streams.filter(abr=selected_abr).first()

            if st.button("Download Audio"):
                if stream:
                    with st.spinner("Downloading audio..."):
                        file_path = stream.download(filename_prefix="audio_")
                    st.success(f"Downloaded audio: {file_path}")
                    with open(file_path, "rb") as f:
                        st.download_button(
                            label="Click here to download audio file",
                            data=f,
                            file_name=f"{yt.title}_{selected_abr}.mp3",
                            mime="audio/mp3"
                        )
                else:
                    st.error("Selected audio bitrate not available.")
