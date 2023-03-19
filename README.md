# Blender Subtitle Editor

A suite of tools to empower working with subtitles/text strips in the Blender Video Sequence Editor. Highlights are multi-language auto-transcription of subtitles, import & export of subtitles, subtitle translation to multiple languages, batch changing of text styles and convenient list editing and navigation of subtitles.

## Tutorial

https://user-images.githubusercontent.com/1322593/223361423-36917ff0-2756-4e80-8c29-ce83096bf085.mp4

## Features
* Import and export of subtitles.
* Transcribe audio to subtitles.
* Translate subtitles.
* List all subtitles in order.
* Edit subtitles in the list.
* Search text across strips. 
* Add and remove subtitles from the list.
* Ripple delete the footage the subtitle line belongs to. 
* Insert line breaks.
* Copy text styling and font from active strip to all selected strips.

## Installation
Download the [zip](https://github.com/tin2tin/subtitle_editor/archive/refs/heads/main.zip) and install.
(If the python modules fail to install, run Blender as administrator, and then run the add-on)

## Where?
In the Sequencer sidebar, the Subtitle Editor tab can be found. Most functions are there, however the function for copying style properties can be found in the Text Strip tab and in the Add menu import Subtitles and Transcription can be found.

## Transcriptions

https://user-images.githubusercontent.com/1322593/223356195-73ffebea-143d-4d36-81eb-c22a8f580bc0.mp4

## Translation

https://user-images.githubusercontent.com/1322593/223357037-8ff2a9c8-9ce3-410d-a344-ec2a64448883.mp4

## Modules used
* [pysubs2](https://github.com/tkarabela/pysubs2)
* [SRTranslator](https://github.com/sinedie/SRTranslator)
* [openai-whisper](https://github.com/openai/whisper)



