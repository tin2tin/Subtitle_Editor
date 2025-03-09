[<img src="https://img.shields.io/badge/Discord%20-%20Invite">](https://discord.gg/HMYpnPzbTm) ![X (formerly Twitter) Follow](https://img.shields.io/twitter/follow/tintwotin)
# Blender Subtitle Editor

A suite of tools to empower working with subtitles/text strips in the Blender Video Sequence Editor. Highlights are multi-language auto-transcription of subtitles, import & export of subtitles, subtitle translation to multiple languages, batch changing of text styles and convenient list editing and navigation of subtitles.

#DES

this is a fork from subtitle editor(https://github.com/tin2tin/Subtitle_Editor), add a new function which can make subtitle to 3d mesh.

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
* Change Whisper model under Preferences.
* System Console pop-up when doing file-related processing.

## Installation
(As for Linux and MacOS, if anything differs in installation, then please share instructions.)
* Use Blender 4.4+
* First you must download and install git for your platform: https://git-scm.com/downloads
* Download the add-on: [zip](https://github.com/darkicerain/subtitle_editor/archive/refs/heads/main.zip)
* On Windows, right click on the Blender icon and "Run Blender as Administrator"(or you'll get write permission errors).
* Install the add-on as usual: Preferences > Add-ons > Install > select file > enable the add-on.
* If the translation feature should be used Firefox needs to be installed.  

Tip           |
:------------- |
If any python modules are missing, use this add-on to manually install them:      |
https://github.com/amb/blender_pip      |

## Where?
In the Sequencer sidebar, the Subtitle Editor tab can be found. Most functions are there, however the function for copying style properties can be found in the Text Strip tab and in the Add menu import Subtitles and Transcription can be found.

## Transcriptions

https://user-images.githubusercontent.com/1322593/223356195-73ffebea-143d-4d36-81eb-c22a8f580bc0.mp4

## Translation

https://user-images.githubusercontent.com/1322593/223357037-8ff2a9c8-9ce3-410d-a344-ec2a64448883.mp4

## Preferences

![image](https://user-images.githubusercontent.com/1322593/229266954-d5397992-7299-4157-b136-47fd3ecc8037.png)

## Modules used
* [pysubs2](https://github.com/tkarabela/pysubs2)
* [SRTranslator](https://github.com/sinedie/SRTranslator)
* [openai-whisper](https://github.com/openai/whisper)



