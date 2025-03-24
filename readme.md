# Yaml-to-PPTX-Video

This tool generates PowerPoint video presentations from a Learn Unit. It takes the [unit].yml file as input and generates a PowerPoint video presentation for the unit.

It does this in the following way:

- Starts a new PowerPoint presentation, based on the template.pptx file.
- Read the [unit].yml file and extracts the uid, title and sections from it. A section is basically a markdown header with the contents of the section.
- For each section, it creates a new slide in the PowerPoint presentation and adds the title to the slide.
- Depending on the type of section, it chooses the slide layout (intro, conclusion, content)
- Using gpt-4o, it generates a summary in the form of a bulleted list for the slide based on the content.
- It also generates a speaker transcript, based on the content.
- The speaker transcript is added to the notes section of the slide
- An audio file is generated and added to the top of the slide (note it is not visible, and does not play automatically)
- A video of an avatar is generated and added to the slide. Both audio and video are exactly the same in terms of speaker transcript.

```sh
source .venv/bin/activate
python3 main.py --yml_file ../learn-pr/wwl-data-ai/fundamentals-machine-learning/1-introduction.yml
```

# pptx-note-to-video

This tool generates videos from the notes section of a PowerPoint presentation. It takes a pptx file as input and generates a video for each slide in the presentation.

# Notebook

This tool can be used to generate video and audio files using a magic command %%audio and %%video.

There is a dependency on ffmpeg, so make sure you have it installed. You can install it using brew on MacOS:

```sh
brew install ffmpeg
```

This is used in the notebook to transform the video, because otherwise the video will not play the audio inline. You can still use the original video.