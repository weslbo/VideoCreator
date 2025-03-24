from __future__ import print_function
from IPython.core.magic import (Magics, magics_class, line_magic, cell_magic, line_cell_magic)
from IPython import get_ipython
from IPython.display import Audio, display, Video
from IPython.display import display, Markdown, Latex, HTML
from urllib.parse import urljoin
import os, fnmatch
import requests
import re
import azure.cognitiveservices.speech as speechsdk
import datetime
import json
import uuid
import time
import subprocess


# The class MUST call this class decorator at creation time
@magics_class
class MyMagics(Magics):
    

    @cell_magic
    def audio(self, line, cell):
        """
        Custom magic command for interacting with Azure OpenAI GPT model.
        Keeps track of conversation history.
        """
        # Wrap the coroutine call using asyncio.run or an event loop
        import nest_asyncio
        import asyncio
        nest_asyncio.apply()
        return asyncio.run(self.audioasync(line, cell))
    
    async def audioasync(self, line, cell):
        service_region = os.getenv("SPEECH_REGION")
        speech_key = os.getenv("SPEECH_KEY")
        speech_config = speechsdk.SpeechConfig(subscription=speech_key, region=service_region)
        speech_config.set_speech_synthesis_output_format(speechsdk.SpeechSynthesisOutputFormat.Audio24Khz96KBitRateMonoMp3)  

        # Use the provided voice or default to "en-US-Ava:DragonHDLatestNeural"
        voice_name = line.strip() or "en-US-AvaMultilingualNeural"
        speech_config.speech_synthesis_voice_name = voice_name

        mp3_filename = f"./audio/{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M')}.mp3"
        file_config = speechsdk.audio.AudioOutputConfig(filename=mp3_filename)
        speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=file_config)  
        result = speech_synthesizer.speak_text_async(cell).get()
    
        display(Audio(mp3_filename, autoplay=True))

    @cell_magic
    def video(self, line, cell):
        """
        Custom magic command for interacting with Azure OpenAI GPT model.
        Keeps track of conversation history.
        """

        # Wrap the coroutine call using asyncio.run or an event loop
        import nest_asyncio
        import asyncio
        nest_asyncio.apply()
        return asyncio.run(self.videoasync(cell))
    
    async def videoasync(self, cell):
        """
        Custom magic command for interacting with Azure OpenAI GPT model.
        Keeps track of conversation history.
        """

        # Wrap the coroutine call using asyncio.run or an event loop
        job_id = str(uuid.uuid4())
        download_url = None

        url = f'https://{os.getenv("SPEECH_REGION")}.api.cognitive.microsoft.com/avatar/batchsyntheses/{job_id}?api-version=2024-04-15-preview'

        header = {
            'Content-Type': 'application/json',
            'Ocp-Apim-Subscription-Key': os.getenv("SPEECH_KEY")
        }

        payload = {
            'synthesisConfig': {
                "voice": 'en-US-AvaMultilingualNeural',
            },
            'customVoices': {
                # "YOUR_CUSTOM_VOICE_NAME": "YOUR_CUSTOM_VOICE_ID"
            },
            "inputKind": "plainText",
            "inputs": [
                {
                    "content": cell,
                },
            ],
            "avatarConfig":
                {
                    "customized": False, # set to True if you want to use customized avatar
                    "talkingAvatarCharacter": 'Lisa',  # talking avatar character
                    "talkingAvatarStyle": 'technical-sitting',  # talking avatar style, required for prebuilt avatar, optional for custom avatar
                    "videoFormat": "mp4",
                    "videoCodec": "h264",
                    "subtitleType": "external_file",
                    "backgroundColor": "#FFFFFFFF", # background color in RGBA format, default is white; can be set to 'transparent' for transparent background
                    #"videoCrop": {  "topLeft": { "x": 560, "y": 0}, "bottomRight": { "x": 1360, "y": 1079}  }
                }  
        }
        
        response = requests.put(url, json.dumps(payload, default=str), headers=header)
        if response.status_code < 400:
            print(f'Job ID: {response.json()["id"]}')
        else:
            print(f'- Failed to submit batch avatar job: [{response.status_code}], {response.text}')

        while True:
            status = MyMagics.get_synthesis(url)
            if status == 'Succeeded':
                print('- batch avatar job succeeded')
                download_url, subtitle_url = MyMagics.getdownloadurl(url)
                
                local_url = f"./video/{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M')}.mp4"
                
                response = requests.get(download_url)
                with open(local_url, 'wb') as file:
                    file.write(response.content) 
                    
                # Define the ffmpeg command
                ffmpeg_command = [
                    "ffmpeg",
                    "-i", local_url,
                    "-c:a", "pcm_s32le",
                    local_url.replace(".mp4", "u.mp4")
                ]

                # Run the ffmpeg command
                subprocess.run(ffmpeg_command, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

                # Display the transformed video
                transformed_video_url = local_url.replace(".mp4", "u.mp4")
                
                display(Video(transformed_video_url))
                
                local_url = f"./video/{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M')}.srt"
                    
                response = requests.get(subtitle_url)
                with open(local_url, 'wb') as file:
                    file.write(response.content)   
                    
                break
            elif status == 'Failed':
                print('- batch avatar job failed')
                break
            else:
                print(f'- batch avatar job is [{status}]')
                time.sleep(5)     
                
    @classmethod
    def get_synthesis(cls, url):
        header = {
            'Ocp-Apim-Subscription-Key': os.getenv("SPEECH_KEY")
        }

        response = requests.get(url, headers=header)
        if response.status_code < 400:
            #print(response.json())
            if response.json()['status'] == 'Succeeded':
                print(f'Download URL: {response.json()["outputs"]["result"]}')
            return response.json()['status']
        else:
            print(f'- Failed to get batch job: {response.text}')       
    
    @classmethod
    def getdownloadurl(cls, url):
        header = {
            'Ocp-Apim-Subscription-Key': os.getenv("SPEECH_KEY")
        }

        response = requests.get(url, headers=header)

        response = requests.get(url, headers=header)
        if response.status_code < 400:
            print('- Get batch avatar job successfully')
            #print(response.json())
            if response.json()['status'] == 'Succeeded':
                return response.json()["outputs"]["result"], response.json()["outputs"]["subtitle"]
        else:
            print(f'- Failed to get batch avatar job: {response.text}')
                
def load_ipython_extension(ipython):
    """
    Any module file that define a function named `load_ipython_extension`
    can be loaded via `%load_ext module.path` or be configured to be
    autoloaded by IPython at startup time.
    """
    # You can register the class itself without instantiating it.  IPython will
    # call the default constructor on it.
    ipython.register_magics(MyMagics)

ip = get_ipython()
load_ipython_extension(ip)