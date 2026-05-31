[<img src="https://img.shields.io/badge/Discord%20-%20Invite">](https://discord.gg/HMYpnPzbTm) ![X (formerly Twitter) Follow](https://img.shields.io/twitter/follow/tintwotin)

# Blender Subtitle Editor

A suite of tools to empower working with subtitles and text strips in the Blender Video Sequence Editor (VSE).

The addon's core feature is high-performance, multi-language audio transcription and translation powered by **Faster Whisper**. It also includes a full set of tools for importing, exporting, editing, and managing subtitles, such as a convenient list editor, batch styling, and timeline navigation.

## Features

### Transcription & Translation
*   **High-Performance Transcription**: Utilizes `Faster Whisper` for fast and accurate audio-to-text conversion.
*   **Translate to English**: Translate audio from any supported language directly into English subtitles.
*   **Extensive Model Support**: Choose from various models (`tiny`, `base`, `small`, `medium`, `large-v1/v2/v3`, and distilled versions) to balance speed and accuracy.
*   **Hardware Acceleration**: Supports CUDA for significant speed-ups on NVIDIA GPUs, as well as CPU execution.
*   **Optimized Performance**: Select compute types like `int8` or `float16` to optimize for speed and memory usage.
*   **Voice Activity Detection (VAD)**: Intelligently filters out silent or non-speech segments for cleaner results.
*   **Direct-to-Timeline**: Automatically creates and places styled text strips in the VSE timeline.

### Editing & Management
*   **List Editor**: View and edit all subtitles in a convenient, ordered list.
*   **Timeline Sync**: Clicking an item in the list jumps to the corresponding strip in the timeline.
*   **Batch Styling**: Copy text styling (font, size, color, shadow, box, etc.) from the active strip to all other selected text strips.
*   **Quick Editing**: Add, remove, or ripple-delete subtitles and their corresponding footage range directly from the list.
*   **Easy Navigation**: Use hotkeys or buttons to jump to the previous/next subtitle in the sequence.
*   **Insert Line Breaks**: Easily add new lines to your subtitle text.

### Import & Export
*   **Broad Format Support**: Import and export subtitles in various formats, including `.srt`, `.ass`, `.vtt`, and `.fountain`.
*   **Legacy Translation**: An option to translate subtitle files on import is available (separate from the main transcription-translation feature).

## Tutorial (List Editor)

This video demonstrates the basics of the subtitle list editor.

https://user-images.githubusercontent.com/1322593/223361423-36917ff0-2756-4e80-8c29-ce83096bf085.mp4

## Installation

1.  **Blender Version**: This addon requires **Blender 5.2** or newer.
2.  **Download**: [Download the add-on as a .zip file](https://github.com/tin2tin/subtitle_editor/archive/refs/heads/main.zip). Do not unzip it.
3.  **Windows Users**: For the initial setup, you must run Blender as an Administrator to allow the addon to install dependencies. Right-click the Blender icon and choose "Run as administrator". This is only needed once.
4.  **Install the Add-on**: In Blender, go to `Edit > Preferences > Add-ons` and click `Install...`. Select the downloaded `.zip` file.
5.  **Enable the Add-on**: Find "Subtitle Editor" in the list and check the box to enable it.
6.  **Install Dependencies**:
    *   Go to the Video Sequencer workspace.
    *   Open the Sidebar (`N` key).
    *   Go to the `Subtitle Editor` tab and find the **Transcription & Translation** panel.
    *   Click the `Install/Verify Dependencies` button.
    *   This will download and install `Faster Whisper` and its required libraries. This may take several minutes. You can monitor the progress in the Blender System Console (`Window > Toggle System Console`).

> [!TIP]
> If the automatic installation fails, you can use the [blender_pip](https://github.com/amb/blender_pip) addon to manually install the required Python modules from the "Modules Used" section below.

## Where to Find It

*   **Main Panels**: In the `Sequencer`, open the UI Sidebar (`N` key) and click the **Subtitle Editor** tab. Here you will find:
    *   **Subtitle Editor**: The list for editing and managing existing text strips.
    *   **Transcription & Translation**: Controls for the `Faster Whisper` integration.
*   **Import**: `Sequencer > Add Menu > Subtitles`.
*   **Export**: From the `Subtitle Editor` list panel, click the export icon.
*   **Copy Styling**: When a Text strip is active, the `Copy Properties to Selected` button appears in the `Sequencer > Sidebar > Strip Tab > Effect` panel.

## Transcription & Translation Panel

All settings for transcription and translation are located directly in the VSE sidebar for easy access.

![image](https://github.com/user-attachments/assets/82594e0a-7af6-4461-a23f-a316b0a5024f)

## Legacy Import Translation

The addon also retains a feature to translate `.srt` files upon import.

https://user-images.githubusercontent.com/1322593/223357037-8ff2a9c8-9ce3-410d-a344-ec2a64448883.mp4

## Modules Used
This add-on relies on the following open-source libraries, which are installed automatically:
*   [Faster Whisper](https://github.com/SYSTRAN/faster-whisper): A reimplementation of OpenAI's Whisper model for faster and more efficient transcription.
*   [pysubs2](https://github.com/tkarabela/pysubs2): A library for editing subtitle files.
