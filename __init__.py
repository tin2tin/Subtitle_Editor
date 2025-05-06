# ##### BEGIN GPL LICENSE BLOCK #####
#
#  This program is free software; you can redistribute it and/or
#  modify it under the terms of the GNU General Public License
#  as published by the Free Software Foundation; either version 2
#  of the License, or (at your option) any later version.
#
#  This program is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.
#
#  You should have received a copy of the GNU General Public License
#  along with this program; if not, write to the Free Software Foundation,
#  Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301, USA.
#
# ##### END GPL LICENSE BLOCK #####

bl_info = {
    "name": "Subtitle Editor",
    "description": "Displays a list of all Subtitle Editor in the VSE and allows editing of their text.",
    "author": "tintwotin",
    "version": (1, 1),
    "blender": (4, 4, 0),
    "location": "Sequencer > Side Bar > Subtitle Editor Tab",
    "warning": "",
    "doc_url": "",
    "support": "COMMUNITY",
    "category": "Sequencer",
}


import os, sys, bpy, pathlib, re, ctypes, site, subprocess, platform
import ensurepip
from bpy.props import (
    EnumProperty,
    StringProperty,
    BoolProperty,
    IntProperty,
    FloatProperty,
    PointerProperty
)
from bpy.types import Operator
from bpy_extras.io_utils import ImportHelper
from datetime import timedelta
os_platform = platform.system()  # 'Linux', 'Darwin', 'Java', 'Windows'

def get_strip_by_name(name):
    for strip in bpy.context.scene.sequence_editor.sequences[0:]:
        if strip.name == name:
            return strip
    return None

FASTER_WHISPER_VERSION = "1.0.3" # Or latest known stable version
REQUIRED_PACKAGE = f"faster-whisper=={FASTER_WHISPER_VERSION}"
dependencies_checked = False
dependencies_installed = False
faster_whisper_module = None


def find_first_empty_channel(start_frame, end_frame):
    for ch in range(1, len(bpy.context.scene.sequence_editor.sequences_all) + 1):
        for seq in bpy.context.scene.sequence_editor.sequences_all:
            if (
                seq.channel == ch
                and seq.frame_final_start < end_frame
                and seq.frame_final_end > start_frame
            ):
                break
        else:
            return ch
    return 1


def update_text(self, context):
    for strip in bpy.context.scene.sequence_editor.sequences:
        if strip.type == "TEXT" and strip.name == self.name:
            # Update the text of the text strip
            strip.text = self.text

            get_strip = get_strip_by_name(self.name)
            if get_strip:
                # Deselect all strips.
                for seq in context.scene.sequence_editor.sequences_all:
                    seq.select = False
                bpy.context.scene.sequence_editor.active_strip = strip
                get_strip.select = True

                # Set the current frame to the start frame of the active strip
                bpy.context.scene.frame_set(int(strip.frame_start))
            break


def show_system_console(show):
    if os_platform == "Windows":
        # https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow
        SW_HIDE = 0
        SW_SHOW = 5

        ctypes.windll.user32.ShowWindow(
            ctypes.windll.kernel32.GetConsoleWindow(), SW_SHOW #if show else SW_HIDE
        )


def set_system_console_topmost(top):
    if os_platform == "Windows":
        # https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowpos
        HWND_NOTOPMOST = -2
        HWND_TOPMOST = -1
        HWND_TOP = 0
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        SWP_NOZORDER = 0x0004

        ctypes.windll.user32.SetWindowPos(
            ctypes.windll.kernel32.GetConsoleWindow(),
            HWND_TOP if top else HWND_NOTOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER,
        )

def get_selected_strip(context):
    """Gets the selected non-meta audio strip."""
    for strip in reversed(context.selected_sequences):
        if strip.type == 'SOUND' and not strip.mute:
            # Check if inside a meta strip and return None if so
            # (This check might need refinement depending on desired behavior with meta strips)
            # if context.selected_sequences and context.selected_sequences[0].type == 'META':
            #     return None
            return strip
    return None

def ensure_user_site_packages(user_site_packages_path):
    """
    Ensures the user site-packages directory exists and attempts to make it writable.
    Returns True on success or if directory exists and is writable, False otherwise.
    """
    if not os.path.exists(user_site_packages_path):
        try:
            os.makedirs(user_site_packages_path, exist_ok=True)
            print(f"Created user site-packages directory: {user_site_packages_path}")
        except OSError as e:
            print(f"Error creating directory {user_site_packages_path}: {e}")
            return False

    if not os.access(user_site_packages_path, os.W_OK):
        print(f"User site-packages directory is not writable: {user_site_packages_path}")
        # Attempting to change permissions might be risky and platform-dependent.
        # On Linux/macOS:
        # try:
        #     os.chmod(user_site_packages_path, stat.S_IRWXU | stat.S_IRGRP | stat.S_IXGRP | stat.S_IROTH | stat.S_IXOTH)
        #     print(f"Attempted to make {user_site_packages_path} writable.")
        #     return os.access(user_site_packages_path, os.W_OK)
        # except Exception as e:
        #      print(f"Could not change permissions for {user_site_packages_path}: {e}")
        #      return False
        # For now, just warn if not writable. Installation might still work if Python has other means.
        return True # Proceed with caution
    return True


def check_faster_whisper():
    """Checks if faster-whisper is installed and importable."""
    global dependencies_installed, faster_whisper_module
    if dependencies_installed and faster_whisper_module:
        return True
    try:
        # Try importing the core component
        from faster_whisper import WhisperModel
        # Store the module for later use if needed (optional)
        import faster_whisper
        faster_whisper_module = faster_whisper
        dependencies_installed = True
        print("faster-whisper found and imported successfully.")
        return True
    except ImportError:
        dependencies_installed = False
        faster_whisper_module = None
        print("faster-whisper module not found.")
        return False
    except Exception as e:
        # Catch other potential import errors
        dependencies_installed = False
        faster_whisper_module = None
        print(f"An unexpected error occurred during faster-whisper import check: {e}")
        return False

def install_dependencies(blender_python_exe):
    """Attempts to install faster-whisper using pip. Returns (bool success, str message)."""
    global dependencies_installed

    # Ensure pip is available in Blender's Python
    try:
        print("Ensuring pip is available...")
        ensurepip.bootstrap()
    except Exception as e:
        print(f"Failed to bootstrap pip: {e}. Will attempt to use pip anyway.")
        # Continue, maybe pip is already there via other means

    # Construct the pip install command
    cmd = [
        blender_python_exe, "-m", "pip", "install", "--upgrade",
        "--user", "--no-cache-dir", "onnxruntime"
    ]
    print(f"Running installation command: {' '.join(cmd)}")

    try:
        # Use subprocess.run with captured output and error checking
        process = subprocess.run(cmd, capture_output=True, text=True, check=True, encoding='utf-8', errors='replace')
    except subprocess.CalledProcessError as e:
        # Pip command failed
        error_message = f"ERROR: Failed to install onnxruntime.\n" \
                        f"Ensure Blender is started as administrator.\n" \
                        f"Command: {' '.join(e.cmd)}\n" \
                        f"Return Code: {e.returncode}\n" \
                        f"--- Error Output (stderr) ---\n{e.stderr}\n" \
                        f"--- Standard Output (stdout) ---\n{e.stdout}\n" \
                        f"-----------------------------"
        print(error_message) # Log detailed error to console
        report_message = f"ERROR: Failed to install onnxruntime. Check Blender System Console for details."
        dependencies_installed = False
        return False, report_message
    
    # Construct the pip install command
    cmd = [
        blender_python_exe, "-m", "pip", "install", "--upgrade",
        "--user", "--no-cache-dir", REQUIRED_PACKAGE
    ]
    print(f"Running installation command: {' '.join(cmd)}")

    try:
        # Use subprocess.run with captured output and error checking
        process = subprocess.run(cmd, capture_output=True, text=True, check=True, encoding='utf-8', errors='replace')
        print("faster-whisper installation command finished.")
        print("--- pip stdout ---")
        print(process.stdout)
        print("--- pip stderr ---")
        print(process.stderr) # Print stderr even on success, might contain warnings
        print("------------------")

        # Verify installation by trying to import again
        if check_faster_whisper():
            dependencies_installed = True
            return True, f"{REQUIRED_PACKAGE} installed successfully!"
        else:
            dependencies_installed = False
            # This is odd, install reported success but import failed
            error_message = f"Installation command seemed to succeed, but '{REQUIRED_PACKAGE}' could not be imported afterwards.\n" \
                            f"Check the Blender System Console for output from pip.\n" \
                            f"Stdout:\n{process.stdout}\nStderr:\n{process.stderr}"
            print(error_message)
            return False, "Installation succeeded but module import failed. Check console & restart Blender."


    except subprocess.CalledProcessError as e:
        # Pip command failed
        error_message = f"ERROR: Failed to install {REQUIRED_PACKAGE}.\n" \
                        f"Command: {' '.join(e.cmd)}\n" \
                        f"Return Code: {e.returncode}\n" \
                        f"--- Error Output (stderr) ---\n{e.stderr}\n" \
                        f"--- Standard Output (stdout) ---\n{e.stdout}\n" \
                        f"-----------------------------"
        print(error_message) # Log detailed error to console
        report_message = f"ERROR: Failed to install {REQUIRED_PACKAGE}. Check Blender System Console for details."
        dependencies_installed = False
        return False, report_message

    except FileNotFoundError:
        error_message = f"ERROR: Blender's Python executable not found at '{blender_python_exe}'. Cannot install dependencies."
        print(error_message)
        dependencies_installed = False
        return False, error_message

    except Exception as e:
        # Catch other potential errors (e.g., permission denied even with --user)
        error_message = f"An unexpected error occurred during installation: {e}"
        import traceback
        traceback.print_exc() # Print traceback to console
        print(error_message)
        dependencies_installed = False
        return False, f"An unexpected error occurred during installation. Check console: {e}"
    


class WhisperProperties(bpy.types.PropertyGroup):
    """Properties for the Faster Whisper Addon"""

    model_size: EnumProperty(
        name="Model Size",
        description="Size of the Faster Whisper model (larger = more accurate, slower, more memory)",
        items=[
            ('tiny', 'Tiny (~39M)', 'Tiny model (~39M params)'),
            ('base', 'Base (~74M)', 'Base model (~74M params)'),
            ('small', 'Small (~244M)', 'Small model (~244M params)'),
            ('medium', 'Medium (~769M)', 'Medium model (~769M params)'),
            ('distil-small.en', 'Distil Small EN (~206M)', 'Distilled Small English-only'),
            ('distil-medium.en', 'Distil Medium EN (~668M)', 'Distilled Medium English-only'),
            ('distil-large-v2', 'Distil Large v2 (~1364M)', 'Distilled Large v2 model'),
            ('large-v1', 'Large v1 (~1550M)', 'Large model v1 (~1550M params)'),
            ('large-v2', 'Large v2 (~1550M)', 'Large model v2 (~1550M params)'),
            ('large-v3', 'Large v3 (~1550M)', 'Large model v3 (~1550M params)'),
        ],
        default='small',
    )

    device: EnumProperty(
        name="Device",
        description="Hardware device (CUDA requires Nvidia GPU and toolkit setup)",
        items=[
            ('cpu', 'CPU', 'Use the Central Processing Unit'),
            ('cuda', 'CUDA', 'Use Nvidia GPU via CUDA'),
            # ('auto', 'Auto', 'Automatically detect best device') # Faster-whisper usually better with explicit
        ],
        default='cpu',
    )

    compute_type: EnumProperty(
        name="Compute Type",
        description="Quantization/precision (int8 = fastest, float32 = most precise)",
        # Check faster-whisper docs for supported types per device/backend
        items=[
            ('int8', 'int8', '8-bit integer (Fastest, low memory, good for CPU)'),
            ('int8_float16', 'int8_float16', 'int8 input, float16 compute (GPU)'),
            # ('int16', 'int16', '16-bit integer'), # Less common? Check docs
            ('float16', 'float16', '16-bit float (Requires capable GPU/CPU)'),
            ('float32', 'float32', '32-bit float (Most precise, slowest, highest memory)'),
            # ('auto', 'Auto', 'Select best based on device') # Usually defaults based on device if not set
        ],
        default='int8', # Good default for CPU
    )

    language: EnumProperty(
        name="Language",
        description="Language spoken in the audio ('auto' to detect)",
        # Items based on Whisper's supported languages - abbreviated list for example
        items=[
            ("auto", "Any language (detect)", ""),
            ("bg", "Bulgarian", "Bulgarian"),
            ("zh", "Chinese", "Chinese"),
            ("cs", "Czech", ""),
            ("da", "Danish", ""),
            ("nl", "Dutch", ""),
            ("en", "English", ""),  # Only usable for source language
            # ("en-US", "English (American)", ""),  # Only usable for destination language
            # ("en-GB", "English (British)", ""),  # Only usable for destination language
            ("et", "Estonian", ""),
            ("fi", "Finnish", ""),
            ("fr", "French", ""),
            ("de", "German", ""),
            ("el", "Greek", ""),
            ("hu", "Hungarian", ""),
            ("id", "Indonesian", ""),
            ("it", "Italian", ""),
            ("ja", "Japanese", ""),
            ("lv", "Latvian", ""),
            ("lt", "Lithuanian", ""),
            ("pl", "Polish", ""),
            ("pt", "Portuguese", ""),  # Only usable for source language
            # ("pt-PT", "Portuguese", ""),  # Only usable for destination language
            # ("pt-BR", "Portuguese (Brazilian)", ""),  # Only usable for destination language
            ("ro", "Romanian", ""),
            ("ru", "Russian", ""),
            ("sk", "Slovak", ""),
            ("sl", "Slovenian", ""),
            ("es", "Spanish", ""),
            ("sv", "Swedish", ""),
            ("tr", "Turkish", ""),
            ("uk", "Ukrainian", ""),
            # Add more languages as needed... See Whisper documentation for full list
        ],
        default='auto',
    )

    beam_size: IntProperty(
        name="Beam Size",
        description="Beam size for decoding (higher can improve accuracy but increases computation)",
        default=5,
        min=1,
    )

    use_vad: BoolProperty(
        name="Use VAD Filter",
        description="Enable Voice Activity Detection (VAD) to filter silence/non-speech",
        default=True,
    )

# --- NEW Properties for Text Strips ---
    output_channel: IntProperty(
        name="Output Channel",
        description="VSE channel to place the subtitle text strips on",
        default=2, # Default to channel 2 (often above main video/audio)
        min=1,
    )

    font_size: IntProperty(
        name="Font Size",
        description="Font size for the subtitle text strips",
        default=50,
        min=1,
        max=1000, # Set a reasonable max
    )

    text_align_y: EnumProperty(
        name="Vertical Align",
        description="Vertical alignment of the text within the frame",
        items=[
            ('TOP', 'Top', 'Align text to the top'),
            ('CENTER', 'Center', 'Align text to the center'),
            ('BOTTOM', 'Bottom', 'Align text to the bottom'),
        ],
        default='BOTTOM',
    )

    wrap_width: FloatProperty(
        name="Wrap Width (0=Off)",
        description="Wrap text width as a factor of frame width (e.g., 0.9 for 90%). 0 disables wrapping",
        default=0.9,
        min=0.0, # 0 disables wrap
        max=1.0,
        subtype='FACTOR',
    )


# --- Operators ---

class SEQUENCER_OT_whisper_setup(Operator):
    """Checks and installs the faster-whisper dependency"""
    bl_idname = "sequencer.whisper_setup"
    bl_label = "Install/Verify Dependencies"
    bl_description = f"Downloads and installs {REQUIRED_PACKAGE} using pip"
    bl_options = {'REGISTER', 'INTERNAL'} # Internal prevents redo panel

    @classmethod
    def poll(cls, context):
        # Allow running anytime to verify or reinstall
        return True

    def execute(self, context):
        global dependencies_checked, dependencies_installed, faster_whisper_module

        # Re-check first, maybe it was installed manually since Blender started
        if not dependencies_installed:
            print("Running initial check before attempting installation...")
            check_faster_whisper()

        if dependencies_installed:
             self.report({'INFO'}, f"{REQUIRED_PACKAGE} seems to be installed.")
             # Maybe add a 'reinstall' option later if needed
             return {'FINISHED'}

        self.report({'INFO'}, "Attempting dependency installation...")

        # 1. Find Blender's Python
        python_exe = sys.executable
        if not python_exe or not os.path.isfile(python_exe):
             python_exe = bpy.app.binary_path_python # Fallback
             if not python_exe or not os.path.isfile(python_exe):
                 self.report({'ERROR'}, "Could not find Blender's Python executable.")
                 return {'CANCELLED'}

        print(f"Using Python executable: {python_exe}")

        # 2. Ensure user site-packages is usable (optional but recommended)
        try:
            user_site_packages = subprocess.check_output(
                [python_exe, "-m", "site", "--user-site"],
                stderr=subprocess.STDOUT # Capture stderr too
            ).strip().decode("utf-8", errors='replace')
            print(f"Targeting user site-packages: {user_site_packages}")
            if not ensure_user_site_packages(user_site_packages):
                 print("Warning: Could not ensure user site-packages directory exists/is writable. Installation might fail.")
                 # self.report({'WARNING'}, "Could not prepare user site-packages dir. Install might fail.")
        except Exception as e:
            print(f"Could not determine user site-packages: {e}. Proceeding anyway.")

        # 3. Attempt installation
        self.report({'INFO'}, f"Installing {REQUIRED_PACKAGE}... This may take a moment. Check console.")
        bpy.context.window_manager.progress_begin(0, 1)
        bpy.context.window_manager.progress_update(0.5) # Indicate activity

        # Trigger a UI update so the message appears
        bpy.context.window_manager.windows.update()

        success, message = install_dependencies(python_exe)

        bpy.context.window_manager.progress_end()

        if success:
            self.report({'INFO'}, message)
        else:
            self.report({'ERROR'}, message)
            # Provide manual installation instructions as a fallback
            manual_instructions = f"Automatic installation failed.\n" \
                                  f"1. Start Blender as Administrator.\n" \
                                  f"2. Restart Blender after installation.\n" \
                                  f"3. Also ensure FFmpeg is installed and in your system PATH."
            print(manual_instructions) # Also print instructions to console
            self.report({'WARNING'}, "See System Console for manual install instructions.")

        # Update state regardless of success/failure for this session
        dependencies_checked = True

        return {'FINISHED'} if success else {'CANCELLED'}


class SEQUENCER_OT_whisper_transcribe(Operator):
    """Transcribes selected audio strip using Faster Whisper into Text Strips""" # Updated description
    bl_idname = "sequencer.whisper_transcribe"
    bl_label = "Transcribe Audio to Text Strips" # Updated label
    bl_options = {'REGISTER', 'UNDO'}

    task: StringProperty(default="transcribe") # "transcribe" or "translate"

    @classmethod
    def poll(cls, context):
        if not context.scene:
            return False
        if not dependencies_installed or faster_whisper_module is None:
             # Disable if not installed (check is cheap)
             cls.poll_message_set("Dependencies not installed. Run 'Install/Verify Dependencies'.")
             return False
        # Check for selected audio strip
        s = get_selected_strip(context)
        if not s:
             cls.poll_message_set("Select an audio strip first.")
             return False
        if not s.sound or not s.sound.filepath:
             cls.poll_message_set("Selected audio strip has no file path.")
             return False
        if not s.type =="SOUND":
             return False
        return True


    def execute(self, context):
        global dependencies_checked, dependencies_installed, faster_whisper_module
        scene = context.scene
        props = scene.whisper_props # Access properties via the property group

        # --- Dependency Check ---
        if not dependencies_checked:
            check_faster_whisper() # Check silently if setup wasn't run
            dependencies_checked = True

        if not dependencies_installed or faster_whisper_module is None:
            self.report({'ERROR'}, f"{REQUIRED_PACKAGE} not installed. Please run '{SEQUENCER_OT_whisper_setup.bl_label}'.")
            # Consider showing a popup:
            # bpy.ops.wm.call_confirm_popup(message=f"Please run '{SEQUENCER_OT_whisper_setup.bl_label}' first.")
            return {'CANCELLED'}

        # --- Get Parameters & Validate ---
        strip = get_selected_strip(context)
        if not strip: # Should be caught by poll, but double check
            self.report({'ERROR'}, "No valid audio strip selected.")
            return {'CANCELLED'}

        try:
            if not strip.sound or not strip.sound.filepath:
                 self.report({'ERROR'}, f"Audio strip '{strip.name}' missing sound data or filepath.")
                 return {'CANCELLED'}

            audio_filepath = bpy.path.abspath(strip.sound.filepath) # Get path once

            if not os.path.exists(audio_filepath): # Check existence once
                self.report({'ERROR'}, f"Audio file not found: {audio_filepath}")
                return {'CANCELLED'}
        except Exception as e:
             self.report({'ERROR'}, f"Error accessing audio filepath for '{strip.name}': {e}")
             import traceback
             traceback.print_exc()
             return {'CANCELLED'}

        allowed_extensions = {'.wav', '.mp3', '.ogg', '.flac', '.m4a', '.aac', '.wma', '.opus'} # Add more if needed
        _ , file_extension = os.path.splitext(audio_filepath)
        if file_extension.lower() not in allowed_extensions:
             # Specifically check for .blend
             if file_extension.lower() == '.blend':
                  error_msg = (f"The source path for the audio strip points to a '.blend' file ({os.path.basename(audio_filepath)}), "
                               f"not an audio file. Please correct the strip's 'File Path' in Blender's properties panel "
                               f"or unpack the audio if it was packed.")
             else:
                  error_msg = (f"The source file ('{os.path.basename(audio_filepath)}') does not appear to be a supported audio format. "
                               f"Supported types include: {', '.join(allowed_extensions)}. "
                               f"Please check the strip's 'File Path' in Blender.")
             self.report({'ERROR'}, error_msg)
             return {'CANCELLED'}

        # Get properties from scene property group
        model_size = props.model_size
        device = props.device
        compute_type = props.compute_type
        language_code = props.language if props.language != "auto" else None # faster-whisper uses None for auto
        beam_size = props.beam_size
        use_vad = props.use_vad
        current_task = self.task # "transcribe" or "translate"

        strip_start_frame = strip.frame_start
        fps = scene.render.fps / scene.render.fps_base
        if fps <= 0:
             self.report({'ERROR'}, "Scene FPS must be positive.")
             return {'CANCELLED'}

        # --- Load Model and Transcribe ---
        try:
            # Access the already imported module's class
            WhisperModel = faster_whisper_module.WhisperModel

            print(f"Loading faster-whisper model: {model_size} (Device: {device}, Compute: {compute_type})")
            self.report({'INFO'}, f"Loading model '{model_size}'... (May download first time)")
            bpy.context.window_manager.windows.update() # Force redraw

            # Check if model exists locally, potentially estimate download size/time? (Advanced)

            model = WhisperModel(model_size, device=device, compute_type=compute_type)

            print(f"Starting transcription...")
            self.report({'INFO'}, f"Transcribing '{os.path.basename(audio_filepath)}' (Task: {current_task})...")
            bpy.context.window_manager.progress_begin(0, 100)

            # Faster-whisper transcribe yields segments
            # VAD default parameters are usually sensible. Can be tuned via vad_parameters=dict(...)
            segments, info = model.transcribe(
                audio=audio_filepath,
                language=language_code,
                task=current_task,
                beam_size=beam_size,
                vad_filter=use_vad,
                vad_parameters=dict(min_silence_duration_ms=500), # Example VAD tuning
                # word_timestamps=False, # Set True if needed, increases computation significantly
                # condition_on_previous_text=True # Helps context, default True
            )

            detected_lang = info.language
            detected_prob = info.language_probability
            print(f"Detected language: {detected_lang} (Confidence: {detected_prob:.2f})")
            print(f"Transcription duration: {info.duration:.2f}s") # Actual processing duration

            if language_code is None and detected_lang:
                 # Try to update the UI language selector if auto-detect was used

                 # --- CORRECTED WAY TO GET LANGUAGE IDENTIFIERS ---
                 # Directly use the list defined in WhisperProperties class definition
                 language_items_definition = [
                    ("auto", "Any language (detect)", ""),
                    ("bg", "Bulgarian", ""),
                    ("zh", "Chinese", ""),
                    ("cs", "Czech", ""),
                    ("da", "Danish", ""),
                    ("nl", "Dutch", ""),
                    ("en", "English", ""),  # Only usable for source language
                    # ("en-US", "English (American)", ""),  # Only usable for destination language
                    # ("en-GB", "English (British)", ""),  # Only usable for destination language
                    ("et", "Estonian", ""),
                    ("fi", "Finnish", ""),
                    ("fr", "French", ""),
                    ("de", "German", ""),
                    ("el", "Greek", ""),
                    ("hu", "Hungarian", ""),
                    ("id", "Indonesian", ""),
                    ("it", "Italian", ""),
                    ("ja", "Japanese", ""),
                    ("lv", "Latvian", ""),
                    ("lt", "Lithuanian", ""),
                    ("pl", "Polish", ""),
                    ("pt", "Portuguese", ""),  # Only usable for source language
                    # ("pt-PT", "Portuguese", ""),  # Only usable for destination language
                    # ("pt-BR", "Portuguese (Brazilian)", ""),  # Only usable for destination language
                    ("ro", "Romanian", ""),
                    ("ru", "Russian", ""),
                    ("sk", "Slovak", ""),
                    ("sl", "Slovenian", ""),
                    ("es", "Spanish", ""),
                    ("sv", "Swedish", ""),
                    ("tr", "Turkish", ""),
                    ("uk", "Ukrainian", ""),
                     # Ensure this list exactly matches the items in WhisperProperties.language
                 ]
                 valid_language_identifiers = {item[0] for item in language_items_definition} # Use a set for efficient lookup
                 # --- END CORRECTION ---

                 if detected_lang in valid_language_identifiers:
                     # Check if update is needed to avoid unnecessary redraws/updates
#                     if props.language != detected_lang:
#                          print(f"Updating UI language from '{props.language}' to detected '{detected_lang}'")
#                          props.language = detected_lang # Update the property group instance
                          # Force a UI redraw if necessary (might not be needed depending on Blender version)
                          # context.area.tag_redraw()
                     self.report({'INFO'}, f"Detected language: {detected_lang}")
                 else:
                     # This case might happen if Whisper detects a language not explicitly listed in our UI
                     print(f"Detected language '{detected_lang}' is not in the predefined EnumProperty items. UI not updated.")


            # --- Process Segments ---
            print("Processing segments...")

            # Get settings for text strips (do this *before* the loop)
            output_channel = props.output_channel
            font_size = props.font_size
            text_align_y = props.text_align_y
            wrap_width = props.wrap_width
            render_width = scene.render.resolution_x

            # Consume the segments generator INTO a list to allow counting and progress
            # Note: This uses more memory for very long transcriptions.
            # An alternative is to iterate directly and estimate progress.
            try:
                segments_list = list(segments)
                num_segments = len(segments_list)
            except Exception as e_consume:
                self.report({'ERROR'}, f"Error processing transcription segments: {e_consume}")
                bpy.context.window_manager.progress_end()
                return {'CANCELLED'}

            if num_segments == 0:
                self.report({'WARNING'}, "No speech segments found in the audio by faster-whisper.")
                bpy.context.window_manager.progress_end()
                return {'FINISHED'} # Exit cleanly if transcription returned nothing

            print(f"Adding {num_segments} text strips to channel {output_channel}...")

            # --- Process Segments and Create Text Strips ---
            created_strips_count = 0
            last_progress_update = -1 # Ensure first update

            # NOW iterate over the POPULATED segments_list
            for i, segment in enumerate(segments_list):
                start_time = segment.start
                end_time = segment.end
                text = segment.text.strip()

                # Calculate frame numbers
                start_frame_calc = round(start_time * fps) + strip_start_frame
                end_frame_calc = round(end_time * fps) + strip_start_frame

                # Explicitly cast to int
                start_frame = int(start_frame_calc)
                end_frame = int(end_frame_calc)
                
                found_channel = find_first_empty_channel(start_frame,end_frame)
                if not output_channel >= found_channel:
                    output_channel = found_channel  

                # Ensure minimum duration of 1 frame
                if end_frame <= start_frame:
                    end_frame = start_frame + 1

                print(f"  {start_time:.2f}s -> {end_time:.2f}s ({start_frame}f -> {end_frame}f): {text}")

                # --- Create the Text Strip ---
                try:
                    text_strip = scene.sequence_editor.sequences.new_effect(
                        name=f"Sub_{start_frame}",
                        type='TEXT',
                        channel=output_channel,
                        frame_start=start_frame,
                        frame_end=end_frame
                    )

                    if text_strip:
                        # Set text content
                        text_strip.text = text

                        # Set appearance properties
                        text_strip.font_size = font_size
                        text_strip.anchor_y = text_align_y
                        text_strip.anchor_x = 'CENTER'

                        # Adjust vertical position slightly
                        if text_align_y == 'BOTTOM':
                            text_strip.location[1] = 0.05
                        elif text_align_y == 'TOP':
                            text_strip.location[1] = 1.0 - 0.05 - (text_strip.font_size / scene.render.resolution_y)
                        else: # CENTER
                             text_strip.location[1] = 0.5

                        # Set wrapping
                        if wrap_width > 0:
                            text_strip.wrap_width = wrap_width
                        else:
                            text_strip.wrap_width = 0

                        # Set other useful defaults
                        text_strip.use_shadow = True
                        text_strip.shadow_color = (0, 0, 0, 1)

                        created_strips_count += 1
                    else:
                         print(f"  ERROR: new_effect call returned None for segment: {text}")

                except Exception as e_strip:
                     print(f"  ERROR creating text strip for segment: {text} -> {e_strip}")
                     import traceback
                     traceback.print_exc()

                # Update progress based on segment index
                progress = int(((i + 1) / num_segments) * 100)
                if progress > last_progress_update:
                    bpy.context.window_manager.progress_update(progress)
                    last_progress_update = progress
                # Allow UI refresh occasionally
                if i % 50 == 0:
                      bpy.context.window_manager.windows.update()

            # End of loop for segments_list

            bpy.context.window_manager.progress_end()
            if created_strips_count > 0:
                # Report success based on actual strips created
                self.report({'INFO'}, f"{current_task.capitalize()} complete. Added {created_strips_count} text strips to channel {output_channel}.")
            elif num_segments > 0:
                # Segments were found by whisper, but strip creation failed for all
                 self.report({'ERROR'}, f"{current_task.capitalize()} finished. Whisper found {num_segments} segments, but failed to create text strips. Check console.")
            else:
                # This case should be caught earlier, but as a fallback
                 self.report({'WARNING'}, f"{current_task.capitalize()} finished, but no segments found or created. Check console.")

        except ImportError:
             # Should be caught by initial check, but safeguard
             self.report({'ERROR'}, "Faster-Whisper module gone missing. Please Setup again or restart Blender.")
             bpy.context.window_manager.progress_end()
             return {'CANCELLED'}
        except Exception as e:
            bpy.context.window_manager.progress_end()
            error_message = f"An error occurred during transcription: {e}"
            print(error_message)
            import traceback
            traceback.print_exc() # Print detailed traceback to Blender System Console
            # Try to provide a more specific common error message
            if "ffmpeg" in str(e).lower():
                 report_msg = "Transcription failed. Ensure FFmpeg is installed and in system PATH. Check Console."
            elif "cuda" in str(e).lower() or "nvtx" in str(e).lower() or "cublas" in str(e).lower():
                 report_msg = "CUDA error. Ensure GPU drivers & CUDA Toolkit are installed correctly. Try CPU device. Check Console."
            elif "memory" in str(e).lower():
                 report_msg = "Out of memory error. Try a smaller model, 'int8' compute type, or increase system RAM/VRAM. Check Console."
            else:
                 report_msg = "Transcription failed. Check Blender System Console for details."
            self.report({'ERROR'}, report_msg)
            return {'CANCELLED'}

        return {'FINISHED'}


def import_module(self, module, install_module):
    show_system_console(True)
    set_system_console_topmost(True)
    module = str(module)

    # Get the path of the Python executable (e.g., python.exe)
    python_exe_dir = os.path.dirname(os.__file__)
    # Construct the path to the site-packages directory
    #site_packages_dir = os.path.join(python_exe_dir, 'lib', 'site-packages')
    site_packages_dir = os.path.join(python_exe_dir, 'lib', 'site-packages') if os.name == 'nt' else os.path.join(python_exe_dir, 'lib', 'python3.x', 'site-packages')        
    # Add the site-packages directory to the top of sys.path
    sys.path.insert(0, site_packages_dir)
    
    #def ensure_pip(self):
    app_path = site.USER_SITE
    if app_path not in sys.path:
        sys.path.append(app_path)
    pybin = sys.executable

    try:
        exec("import " + module)
    except ModuleNotFoundError:
        app_path = site.USER_SITE
        if app_path not in sys.path:
            sys.path.append(app_path)
        pybin = sys.executable

        self.report({"INFO"}, "Installing: " + module + " module.")
        print("Installing: " + module + " module")
        subprocess.call([pybin, "-m", "pip", "install", install_module, "--no-warn-script-location"]) #"--user",  , '--target', site_packages_dir
        try:
            exec("import " + module)
        except ModuleNotFoundError:
            self.report({"INFO"}, "Not found: " + module + " module.")
            print("Not found: " + module + " module")            
            return False
    #show_system_console(True)
    #set_system_console_topmost(False)
    return True


# Define a custom property group to hold the text strip name and text
class TextStripItem(bpy.types.PropertyGroup):
    name: bpy.props.StringProperty()
    text: bpy.props.StringProperty(update=update_text, options={"TEXTEDIT_UPDATE"})
    selected: bpy.props.IntProperty()


class SEQUENCER_UL_List(bpy.types.UIList):
    def draw_item(
        self, context, layout, data, item, icon, active_data, active_propname
    ):
        layout.prop(item, "text", text="", emboss=False)


# Define an operator to refresh the list
class SEQUENCER_OT_refresh_list(bpy.types.Operator):
    """Sync items in the list with the text strips"""

    bl_idname = "text.refresh_list"
    bl_label = "Refresh List"


    def execute(self, context):
        
        active = context.scene.sequence_editor.active_strip
        # Clear the list
        context.scene.text_strip_items.clear()

        # Get a list of all Text Strips in the VSE
        text_strips = [
            strip
            for strip in bpy.context.scene.sequence_editor.sequences
            if strip.type == "TEXT"
        ]

        # Sort the Subtitle Editor based on their start times in the timeline
        text_strips.sort(key=lambda strip: strip.frame_start)

        # Iterate through the sorted text strips and add them to the list
        for strip in text_strips:
            item = context.scene.text_strip_items.add()
            item.name = strip.name
            item.text = strip.text
        # Select only the active strip in the UI list
        for seq in context.scene.sequence_editor.sequences_all:
            seq.select = False
        if active:
            context.scene.sequence_editor.active_strip = active
            active.select = True
            # Set the current frame to the start frame of the active strip
            bpy.context.scene.frame_set(int(active.frame_start))
        return {"FINISHED"}


class SEQUENCER_OT_add_strip(bpy.types.Operator):
    """Add a new text strip after the position of the current selected list item"""

    bl_idname = "text.add_strip"
    bl_label = "Add Text Strip"
    bl_options = {"REGISTER", "UNDO"}

    def execute(self, context):
        scene = context.scene
        text_strips = scene.sequence_editor.sequences_all

        # Get the selected text strip from the UI list
        index = scene.text_strip_items_index
        items = context.scene.text_strip_items
        un_selected = len(items) - 1 < index
        if not un_selected:
            strip_name = items[index].name
            strip = get_strip_by_name(strip_name)
            cf = context.scene.frame_current
            in_frame = cf + strip.frame_final_duration
            out_frame = cf + (2 * strip.frame_final_duration)
            chan = find_first_empty_channel(in_frame, out_frame)

            # Add a new text strip after the selected strip
            strips = scene.sequence_editor.sequences
            new_strip = strips.new_effect(
                name="Text",
                type="TEXT",
                channel=chan,
                frame_start=in_frame,
                frame_end=out_frame,
            )

            # Copy the settings
            if strip and new_strip:
                new_strip.text = "Text"
                new_strip.font_size = strip.font_size
                new_strip.font = strip.font
                new_strip.color = strip.color
                new_strip.use_shadow = strip.use_shadow
                new_strip.blend_type = strip.blend_type
                new_strip.use_bold = strip.use_bold
                new_strip.use_italic = strip.use_italic
                new_strip.shadow_color = strip.shadow_color
                new_strip.use_box = strip.use_box
                new_strip.box_margin = strip.box_margin
                new_strip.location = strip.location
                new_strip.anchor_x = strip.anchor_x
                new_strip.anchor_y = strip.anchor_y
                context.scene.sequence_editor.active_strip = new_strip
            self.report({"INFO"}, "Copying settings from the selected item")
        else:
            strips = scene.sequence_editor.sequences
            chan = find_first_empty_channel(
                context.scene.frame_current, context.scene.frame_current + 100
            )

            new_strip = strips.new_effect(
                name="Text",
                type="TEXT",
                channel=chan,
                frame_start=context.scene.frame_current,
                frame_end=context.scene.frame_current + 100,
            )
            new_strip.wrap_width = 0.68
            new_strip.font_size = 44
            new_strip.location[1] = 0.25
            new_strip.anchor_x = "CENTER"
            new_strip.anchor_y = "TOP"
            new_strip.use_shadow = True
            new_strip.use_box = True
            context.scene.sequence_editor.active_strip = new_strip
            new_strip.select = True
        # Refresh the UIList
        bpy.ops.text.refresh_list()

        # Select the new item in the UIList
        context.scene.text_strip_items_index = index + 1

        return {"FINISHED"}


class SEQUENCER_OT_delete_strip(bpy.types.Operator):
    """Remove item and ripple delete within its range"""

    bl_idname = "text.delete_strip"
    bl_label = "Remove Item and Ripple Delete"

    @classmethod
    def poll(cls, context):
        index = context.scene.text_strip_items_index
        items = context.scene.text_strip_items
        selected = len(items) - 1 < index
        return not selected and context.scene.sequence_editor is not None

    def execute(self, context):
        scene = context.scene
        seq_editor = scene.sequence_editor
        frame_org = scene.frame_current
        index = context.scene.text_strip_items_index
        items = context.scene.text_strip_items
        if len(items) - 1 < index:
            return {"CANCELLED"}
        # Get the selected text strip from the UI list
        strip_name = items[index].name
        strip = get_strip_by_name(strip_name)
        if strip:
            # Select all strips
            for seq in bpy.context.scene.sequence_editor.sequences_all:
                seq.select = True
            # Split out
            scene.frame_current = strip.frame_final_start + strip.frame_final_duration
            bpy.ops.sequencer.split(frame=scene.frame_current, type="SOFT", side="LEFT")

            # Select all strips
            for seq in bpy.context.scene.sequence_editor.sequences_all:
                seq.select = True
            # Split in
            frame_current = scene.frame_current = strip.frame_final_start
            bpy.ops.sequencer.split(frame=frame_current, type="SOFT", side="RIGHT")

            # Deslect all
            for s in bpy.context.scene.sequence_editor.sequences_all:
                s.select = False
            # Select all between in and out
            frame_range = range(
                strip.frame_final_start + 1,
                strip.frame_final_start + strip.frame_final_duration,
            )
            for s in frame_range:
                # scene.frame_current = s
                for stri in bpy.context.scene.sequence_editor.sequences_all:
                    if (
                        stri.frame_final_start
                        <= s
                        <= stri.frame_final_start + stri.frame_final_duration
                    ):
                        stri.select = True
            scene.frame_current = strip.frame_final_start + 1
            bpy.ops.sequencer.delete()
            bpy.ops.sequencer.gap_remove(all=True)

            # Remove the UI list item
            items.remove(index)
        # Refresh the UIList
        bpy.ops.text.refresh_list()

        scene.frame_current = frame_org

        return {"FINISHED"}


class SEQUENCER_OT_delete_item(bpy.types.Operator):
    """Remove strip and item from the list"""

    bl_idname = "text.delete_item"
    bl_label = "Remove Item and Strip"

    @classmethod
    def poll(cls, context):
        index = context.scene.text_strip_items_index
        items = context.scene.text_strip_items
        selected = (len(items) - 1) < index
        return not selected and context.scene.sequence_editor is not None

    def execute(self, context):
        scene = context.scene
        seq_editor = scene.sequence_editor
        index = context.scene.text_strip_items_index
        items = context.scene.text_strip_items
        if len(items) - 1 < index:
            return {"CANCELLED"}
        # Get the selected text strip from the UI list
        strip_name = items[index].name
        strip = get_strip_by_name(strip_name)
        if strip:
            # Deselect all strips
            for seq in context.scene.sequence_editor.sequences_all:
                seq.select = False
            # Delete the strip
            strip.select = True
            bpy.ops.sequencer.delete()

            # Remove the UI list item
            items.remove(index)
        # Refresh the UIList
        bpy.ops.text.refresh_list()

        return {"FINISHED"}


class SEQUENCER_OT_select_next(bpy.types.Operator):
    """Select the item below"""

    bl_idname = "text.select_next"
    bl_label = "Select Next"

    def execute(self, context):
        scene = context.scene
        current_index = context.scene.text_strip_items_index
        max_index = len(context.scene.text_strip_items) - 1

        if current_index < max_index:
            context.scene.text_strip_items_index += 1
            selected_item = scene.text_strip_items[scene.text_strip_items_index]
            selected_item.select = True
            update_text(selected_item, context)
        return {"FINISHED"}


class SEQUENCER_OT_select_previous(bpy.types.Operator):
    """Select the item above"""

    bl_idname = "text.select_previous"
    bl_label = "Select Previous"

    def execute(self, context):
        scene = context.scene
        current_index = context.scene.text_strip_items_index
        if current_index > 0:
            context.scene.text_strip_items_index -= 1
            selected_item = scene.text_strip_items[scene.text_strip_items_index]
            print(selected_item)
            selected_item.select = True
            update_text(selected_item, context)
        return {"FINISHED"}


class SEQUENCER_OT_insert_newline(bpy.types.Operator):
    bl_idname = "text.insert_newline"
    bl_label = "Insert Newline"
    bl_description = "Inserts a newline character at the cursor position"
    bl_options = {"REGISTER", "UNDO"}

    @classmethod
    def poll(cls, context):
        index = context.scene.text_strip_items_index
        items = context.scene.text_strip_items
        selected = len(items) - 1 < index
        return not selected and context.scene.sequence_editor is not None

    def execute(self, context):
        context.scene.text_strip_items[
            context.scene.text_strip_items_index
        ].text += chr(10)
        self.report(
            {"INFO"}, "New line character inserted in the end of the selected item"
        )
        return {"FINISHED"}

def load_lyrics(self, file, context, offset):
    app_path = site.USER_SITE
    if app_path not in sys.path:
        sys.path.append(app_path)
    pybin = sys.executable
    
    print("Please wait. Checking pylrc module...")

    if not import_module(self, "pylrc", "pylrc"):
        return {"CANCELLED"}
    
    import pylrc

    render = bpy.context.scene.render
    fps = render.fps / render.fps_base
    fps_conv = fps #/ 1000

    editor = bpy.context.scene.sequence_editor
    sequences = bpy.context.sequences
    if not sequences:
        addSceneChannel = 1
    else:
        channels = [s.channel for s in sequences]
        channels = sorted(list(set(channels)))
        empty_channel = channels[-1] + 1
        addSceneChannel = empty_channel

    subs = None
    with open(file, "r", encoding="UTF-8") as lyric_file:
        subs = pylrc.parse(lyric_file.read())

    if not subs:
        print("No file imported.")
        self.report({"INFO"}, "No file imported")
        return {"CANCELLED"}

    for i in range(len(subs)):
        line = subs[i]
        print(str(line.time) + ":" + line.text)
        line.start = line.time
        if (i < len(subs) - 1):
            line.end = subs[i + 1].time
        else:
            line.end = line.start + 100

        if line.end and line.text and line.start:
            new_strip = editor.sequences.new_effect(
                name=line.text,
                type="TEXT",
                channel=addSceneChannel,
                frame_start=int(line.start * fps_conv) + offset,
                frame_end=int(line.end * fps_conv) + offset,
            )
            new_strip.text = line.text
            new_strip.wrap_width = 0.68
            new_strip.font_size = 44
            new_strip.location[1] = 0.25
            new_strip.anchor_x = "CENTER"
            new_strip.anchor_y = "TOP"
            new_strip.use_shadow = True
            new_strip.use_box = True
    # Refresh the UIList
    bpy.ops.text.refresh_list()
    
 

def load_subtitles(self, file, context, offset):
    
    #def ensure_pip(self):
    app_path = site.USER_SITE
    if app_path not in sys.path:
        sys.path.append(app_path)
    pybin = sys.executable
    
    print("Please wait. Checking pysubs2 module...")

    if not import_module(self, "pysubs2", "pysubs2"):
        return {"CANCELLED"}
    
    import pysubs2

    render = bpy.context.scene.render
    fps = render.fps / render.fps_base
    fps_conv = fps / 1000

    editor = bpy.context.scene.sequence_editor
    sequences = bpy.context.sequences
    if not sequences:
        addSceneChannel = 1
    else:
        channels = [s.channel for s in sequences]
        channels = sorted(list(set(channels)))
        empty_channel = channels[-1] + 1
        addSceneChannel = empty_channel
    if pathlib.Path(file).is_file():
        if (
            pathlib.Path(file).suffix
            not in pysubs2.formats.FILE_EXTENSION_TO_FORMAT_IDENTIFIER
        ):
            print("Unable to extract subtitles from file")
            self.report({"INFO"}, "Unable to extract subtitles from file")
            return {"CANCELLED"}
    #try:
    subs = pysubs2.load(file, fps=fps, encoding="utf-8")
#    except:
#        print("Import failed. Text encoding must be in UTF-8.")
#        self.report({"INFO"}, "Import failed. Text encoding must be in UTF-8.")
#        return {"CANCELLED"}
    if not subs:
        print("No file imported.")
        self.report({"INFO"}, "No file imported")
        return {"CANCELLED"}
    for line in subs:
        italics = False
        bold = False
        position = False
        pos = []
        line.text = line.text.replace("\\N", "\n")
        print(line.text)
        print(line.start)
        print(line.end)
        
        if line.start == line.end or not line.end:
            line.end = line.start + 100
        if r"<i>" in line.text or r"{\i1}" in line.text or r"{\i0}" in line.text:
            italics = True
            line.text = line.text.replace("<i>", "")
            line.text = line.text.replace("</i>", "")
            line.text = line.text.replace("{\\i0}", "")
            line.text = line.text.replace("{\\i1}", "")
        if r"<b>" in line.text or r"{\b1}" in line.text or r"{\b0}" in line.text:
            bold = True
            line.text = line.text.replace("<b>", "")
            line.text = line.text.replace("</b>", "")
            line.text = line.text.replace("{\\b0}", "")
            line.text = line.text.replace("{\\b1}", "")
        if r"{" in line.text:
            pos_trim = re.search(r"\{\\pos\((.+?)\)\}", line.text)
            pos_trim = pos_trim.group(1)
            pos = pos_trim.split(",")
            x = int(pos[0]) / render.resolution_x
            y = (render.resolution_y - int(pos[1])) / render.resolution_y
            position = True
            line.text = re.sub(r"{.+?}", "", line.text)
        if line.end and line.text and line.start:
            new_strip = editor.sequences.new_effect(
                name=line.text,
                type="TEXT",
                channel=addSceneChannel,
                frame_start=int(line.start * fps_conv) + offset,
                frame_end=int(line.end * fps_conv) + offset,
            )
            new_strip.text = line.text
            new_strip.wrap_width = 0.68
            new_strip.font_size = 44
            new_strip.location[1] = 0.25
            new_strip.anchor_x = "CENTER"
            new_strip.anchor_y = "TOP"
            new_strip.use_shadow = True
            new_strip.use_box = True
            if position:
                new_strip.location[0] = x
                new_strip.location[1] = y
                new_strip.anchor_x = "LEFT"
            if italics:
                new_strip.use_italic = True
            if bold:
                new_strip.use_bold = True
    # Refresh the UIList
    bpy.ops.text.refresh_list()


def check_overlap(strip1, start, end):
    # Check if the strips overlap.
    # print(str(strip1.frame_final_start + strip1.frame_final_duration)+">="+str(start)+" "+str(strip1.frame_final_start) +" <= "+str(end))
    return (strip1.frame_final_start + strip1.frame_final_duration) >= int(start) and (
        strip1.frame_final_start
    ) <= int(end)


# Define the options for the enum
load_models = [
    ("TINY", "Tiny", "Use the tiny model"),
    ("BASE", "Base", "Use the base model"),
    ("SMALL", "Small", "Use the small model"),
    ("MEDIUM", "Medium", "Use the medium model"),
    ("LARGE", "Large", "Use the large model"),
    ("LARGE-V2", "Large v.2", "Use the large v2 model"),
    ("LARGE-V3", "Large v.3", "Use the large v3 model"),
]

# Define the preferences class
class subtitle_preferences(bpy.types.AddonPreferences):
    bl_idname = __name__
    
    # Add the enum option to the preferences
    load_model: bpy.props.EnumProperty(
        name="Model",
        description="Choose the model. And expect higher load times.",
        items=load_models,
        default="LARGE",
    )
    
    def draw(self, context):
        layout = self.layout
        layout.prop(self, "load_model")


#def format_srt_time(seconds):
#    """Convert precise seconds (float) to SRT format HH:MM:SS,mmm"""
#    milliseconds = int(seconds * 1000)  # Convert to milliseconds
#    td = timedelta(milliseconds=milliseconds)
#    return str(td)[:-3].replace(".", ",")  # Format as HH:MM:SS,mmm


def format_srt_time(ms):
    """Convert milliseconds to SRT format HH:MM:SS,mmm"""
    td = timedelta(milliseconds=int(ms))
    return f"{int(td.total_seconds() // 3600):02}:{int((td.total_seconds() % 3600) // 60):02}:{int(td.total_seconds() % 60):02},{int(ms % 1000):03}"


def add_punctuation(text):
    # Basic example: Add a period at the end if it doesn't have one
    if text and text[-1] not in ['.', '!', '?']:
        text += '.'
    return text


class TEXT_OT_transcribe(bpy.types.Operator):
    bl_idname = "text.transcribe"
    bl_label = "Transcription"
    bl_description = "Transcribe audiofile to text strips"
    bl_options = {"REGISTER", "UNDO"}

    @classmethod
    def poll(cls, context):
        return context.scene.sequence_editor

    # Supported languages
    # ['Auto detection', 'Afrikaans', 'Albanian', 'Amharic', 'Arabic', 'Armenian', 'Assamese', 'Azerbaijani', 'Bashkir', 'Basque', 'Belarusian', 'Bengali', 'Bosnian', 'Breton', 'Bulgarian', 'Burmese', 'Castilian', 'Catalan', 'Chinese', 'Croatian', 'Czech', 'Danish', 'Dutch', 'English', 'Estonian', 'Faroese', 'Finnish', 'Flemish', 'French', 'Galician', 'Georgian', 'German', 'Greek', 'Gujarati', 'Haitian', 'Haitian Creole', 'Hausa', 'Hawaiian', 'Hebrew', 'Hindi', 'Hungarian', 'Icelandic', 'Indonesian', 'Italian', 'Japanese', 'Javanese', 'Kannada', 'Kazakh', 'Khmer', 'Korean', 'Lao', 'Latin', 'Latvian', 'Letzeburgesch', 'Lingala', 'Lithuanian', 'Luxembourgish', 'Macedonian', 'Malagasy', 'Malay', 'Malayalam', 'Maltese', 'Maori', 'Marathi', 'Moldavian', 'Moldovan', 'Mongolian', 'Myanmar', 'Nepali', 'Norwegian', 'Nynorsk', 'Occitan', 'Panjabi', 'Pashto', 'Persian', 'Polish', 'Portuguese', 'Punjabi', 'Pushto', 'Romanian', 'Russian', 'Sanskrit', 'Serbian', 'Shona', 'Sindhi', 'Sinhala', 'Sinhalese', 'Slovak', 'Slovenian', 'Somali', 'Spanish', 'Sundanese', 'Swahili', 'Swedish', 'Tagalog', 'Tajik', 'Tamil', 'Tatar', 'Telugu', 'Thai', 'Tibetan', 'Turkish', 'Turkmen', 'Ukrainian', 'Urdu', 'Uzbek', 'Valencian', 'Vietnamese', 'Welsh', 'Yiddish', 'Yoruba']

    def execute(self, context):
        
#        #def ensure_pip(self):
#        app_path = site.USER_SITE
#        if app_path not in sys.path:
#            sys.path.append(app_path)
        pybin = sys.executable
        print("pybin: "+str(pybin))

        # Get the path of the Python executable (e.g., python.exe)
        python_exe_dir = os.path.dirname(os.__file__)
        # Construct the path to the site-packages directory
        site_packages_dir = os.path.join(python_exe_dir, 'lib', 'site-packages') if os.name == 'nt' else os.path.join(python_exe_dir, 'lib', 'python3.x', 'site-packages')        

        # Add the site-packages directory to the top of sys.path
        sys.path.insert(0, site_packages_dir)
        print("site_packages_dir: "+str(site_packages_dir))

#        print("Ensuring: pip")
#        try:
#            subprocess.call([pybin, "-m", "ensurepip"])
#            #subprocess.call([pybin, "-m", "pip", "install", "--upgrade","pip"])
#        except ImportError:
#            print("Pip installation failed!")
#            #return False
#        #return True        
        
        
        print("Please wait. Checking torch & whisper modules...")
        #import_module(self, "torch", "torch==2.0.0")

        if os_platform == "Windows":
            subprocess.call(
                [
                    pybin,
                    "-m",
                    "pip",
                    "install",
                    "torch",
                    "--index-url",
                    "https://download.pytorch.org/whl/cu124",
                    "--no-warn-script-location",
                    #"--user",
                    #'--target', site_packages_dir,
                ]
            )
        else:
             subprocess.call(
                [
                    pybin,
                    "-m",
                    "pip",
                    "install",
                    "torch",
                    "--no-warn-script-location",
                    #"--user",
                    #'--target', site_packages_dir,
                ]
            )  
                 
        if os_platform == "Windows":
            try:
                exec("import triton")
            except ModuleNotFoundError:
                subprocess.call(
                    [
                        pybin,
                        "-m",
                        "pip",
                        "install",
                        "--disable-pip-version-check",
                        "--use-deprecated=legacy-resolver",
                        "triton-windows",
                        #"https://github.com/woct0rdho/triton-windows/releases/download/v3.2.0-windows.post9/triton-3.2.0-cp311-cp311-win_amd64.whl",
                        #"https://hf-mirror.com/LightningJay/triton-2.1.0-python3.11-win_amd64-wheel/resolve/main/triton-2.1.0-cp311-cp311-win_amd64.whl",
                        "--no-warn-script-location",
                        "--upgrade",
                        #'--target', site_packages_dir,
                    ]
                )                        
        else:
            try:
                exec("import triton")
            except ModuleNotFoundError:
                import_module(self, "triton", "triton")

        import_module(self, "whisper", "openai-whisper") # "openai_whisper"):#
        #import_module(self, "whisper", "git+https://github.com/openai/whisper.git") # "openai_whisper"):#

        print("Checking srttranslator module...")
        if not import_module(self, "srtranslator", "srtranslator"):
            print("Srttranslator module not found!")
            return {"CANCELLED"}

        from srtranslator import SrtFile

        import whisper
        from whisper.utils import get_writer
        current_scene = bpy.context.scene
        try:
            active = current_scene.sequence_editor.active_strip
        except AttributeError:
            self.report({"INFO"}, "No active strip selected!")
            return {"CANCELLED"}
        if not active:
            self.report({"INFO"}, "No active strip!")
            return {"CANCELLED"}
        if not active.type == "SOUND":
            self.report({"INFO"}, "Active strip is not a sound strip!")
            return {"CANCELLED"}
        offset = int(active.frame_start)
        clip_start = int(active.frame_start + active.frame_offset_start)
        clip_end = int(
            active.frame_start + active.frame_final_duration
        ) 
        sound_path = bpy.path.abspath(active.sound.filepath)
        sound_path = os.path.normpath(os.path.realpath(sound_path))  # Fully resolve path
        output_dir = os.path.dirname(sound_path)
        print("sound_path: "+str(sound_path))
        print("output_dir: "+str(output_dir))
        audio_basename = os.path.basename(sound_path)
        print("audio_basename: "+str(audio_basename))

        print("Please wait. Processing file...")
        load_model = context.preferences.addons[__name__].preferences.load_model
        model = whisper.load_model(load_model.lower())

        out_dir = os.path.join(output_dir, audio_basename + ".srt")
        if os.path.exists(out_dir):
            os.remove(out_dir)

        transcribe = model.transcribe(sound_path, word_timestamps=True)
        segments = transcribe["segments"]

        silence_threshold = 1000  # 1 second of silence before breaking subtitles

        with open(out_dir, "w", encoding="utf-8") as srtFile:
            segmentId = 0
            last_end_time = 0

            for segment in segments:
                words = segment.get("words", [])
                if not words:
                    continue  # Skip empty segments

                start_time = words[0]["start"] * 1000  # Convert to ms
                end_time = words[-1]["end"] * 1000  # Convert to ms

                # If there's a big silence gap before this segment, adjust timing
                if start_time - last_end_time > silence_threshold:
                    start_time = last_end_time  # Fix start time

                last_end_time = end_time  # Update end time

                startTime = format_srt_time(start_time)
                endTime = format_srt_time(end_time)

                text = segment["text"].strip()
                if not text:
                    continue  # Skip empty text
                text = add_punctuation(text)

                srt_segment = f"{segmentId + 1}\n{startTime} --> {endTime}\n{text}\n\n"
                srtFile.write(srt_segment)

                segmentId += 1

        print("Processing finished.")
        if os.path.exists(out_dir):
            load_subtitles(self, out_dir, context, offset)
        return {"FINISHED"}


class SEQUENCER_OT_import_subtitles(Operator, ImportHelper):
    """Import Subtitles"""

    bl_idname = "sequencer.import_subtitles"
    bl_label = "Import"
    bl_options = {"REGISTER", "UNDO"}

    filename_ext = [".srt", ".ass"]

    filter_glob: StringProperty(
        default="*.srt;*.ass;*.ssa;*.mpl2;*.tmp;*.vtt;*.microdvd;*.lrc",
        options={"HIDDEN"},
        maxlen=255,
    )

    do_translate: BoolProperty(
        name="Translate Subtitles",
        description="Translate subtitles",
        default=False,
        options={"HIDDEN"},
    )

    translate_from: EnumProperty(
        name="From",
        description="Translate from",
        items=(
            ("auto", "Any language (detect)", ""),
            ("bg", "Bulgarian", ""),
            ("zh", "Chinese", ""),
            ("cs", "Czech", ""),
            ("da", "Danish", ""),
            ("nl", "Dutch", ""),
            ("en", "English", ""),  # Only usable for source language
            # ("en-US", "English (American)", ""),  # Only usable for destination language
            # ("en-GB", "English (British)", ""),  # Only usable for destination language
            ("et", "Estonian", ""),
            ("fi", "Finnish", ""),
            ("fr", "French", ""),
            ("de", "German", ""),
            ("el", "Greek", ""),
            ("hu", "Hungarian", ""),
            ("id", "Indonesian", ""),
            ("it", "Italian", ""),
            ("ja", "Japanese", ""),
            ("lv", "Latvian", ""),
            ("lt", "Lithuanian", ""),
            ("pl", "Polish", ""),
            ("pt", "Portuguese", ""),  # Only usable for source language
            # ("pt-PT", "Portuguese", ""),  # Only usable for destination language
            # ("pt-BR", "Portuguese (Brazilian)", ""),  # Only usable for destination language
            ("ro", "Romanian", ""),
            ("ru", "Russian", ""),
            ("sk", "Slovak", ""),
            ("sl", "Slovenian", ""),
            ("es", "Spanish", ""),
            ("sv", "Swedish", ""),
            ("tr", "Turkish", ""),
            ("uk", "Ukrainian", ""),
        ),
        default="auto",
        options={"HIDDEN"},
    )

    translate_to: EnumProperty(
        name="To",
        description="Translate to",
        items=(
            # ("auto", "Any language (detect)", ""),
            ("bg", "Bulgarian", ""),
            ("zh", "Chinese", ""),
            ("cs", "Czech", ""),
            ("da", "Danish", ""),
            ("nl", "Dutch", ""),
            # ("en", "English", ""),  # Only usable for source language
            ("en-US", "English (American)", ""),  # Only usable for destination language
            ("en-GB", "English (British)", ""),  # Only usable for destination language
            ("et", "Estonian", ""),
            ("fi", "Finnish", ""),
            ("fr", "French", ""),
            ("de", "German", ""),
            ("el", "Greek", ""),
            ("hu", "Hungarian", ""),
            ("id", "Indonesian", ""),
            ("it", "Italian", ""),
            ("ja", "Japanese", ""),
            ("lv", "Latvian", ""),
            ("lt", "Lithuanian", ""),
            ("pl", "Polish", ""),
            # ("pt", "Portuguese", ""),  # Only usable for source language
            ("pt-PT", "Portuguese", ""),  # Only usable for destination language
            (
                "pt-BR",
                "Portuguese (Brazilian)",
                "",
            ),  # Only usable for destination language
            ("ro", "Romanian", ""),
            ("ru", "Russian", ""),
            ("sk", "Slovak", ""),
            ("sl", "Slovenian", ""),
            ("es", "Spanish", ""),
            ("sv", "Swedish", ""),
            ("tr", "Turkish", ""),
            ("uk", "Ukrainian", ""),
        ),
        default="en-US",
        options={"HIDDEN"},
    )

    def execute(self, context):
        if self.do_translate:
            #ensure_pip(self)
            print("Please wait. Checking srttranslator module...")
            if not import_module(self, "srtranslator", "srtranslator"):
                return {"CANCELLED"}
            file = self.filepath
            if not file:
                return {"CANCELLED"}
                print("Invalid file")
                self.report({"INFO"}, "Invalid file")
            from srtranslator import SrtFile
            from srtranslator.translators.deepl_scrap import DeeplTranslator
            translator = DeeplTranslator()
            if pathlib.Path(file).is_file():
                print("Translating. Please Wait.")
                self.report({"INFO"}, "Translating. Please Wait.")
                srt = SrtFile(file)
                srt.translate(translator, self.translate_from, self.translate_to)

                # Making the result subtitles prettier
                srt.wrap_lines()

                srt.save(f"{os.path.splitext(file)[0]}_translated.srt")
                translator.quit()
                print("Translating finished.")
                self.report({"INFO"}, "Translating finished.")
        file = self.filepath
        if self.do_translate:
            file = f"{os.path.splitext(file)[0]}_translated.srt"
        if not file:
            return {"CANCELLED"}
        offset = 0
        if file.endswith(".lrc"):
            load_lyrics(self, file, context, offset)
        else:
            load_subtitles(self, file, context, offset)

        return {"FINISHED"}

    def draw(self, context):
        pass


class SEQUENCER_PT_import_subtitles(bpy.types.Panel):
    bl_space_type = "FILE_BROWSER"
    bl_region_type = "TOOL_PROPS"
    bl_label = "Translation"
    bl_parent_id = "FILE_PT_operator"

    @classmethod
    def poll(cls, context):
        sfile = context.space_data
        operator = sfile.active_operator
        return operator.bl_idname == "SEQUENCER_OT_import_subtitles"

    def draw(self, context):
        layout = self.layout
        layout.use_property_split = True
        layout.use_property_decorate = False  # No animation.

        sfile = context.space_data
        operator = sfile.active_operator
        layout = layout.column(heading="Translate Subtitles")
        layout.prop(operator, "do_translate", text="")
        col = layout.column(align=False)
        col.prop(operator, "translate_from", text="From")
        col.prop(operator, "translate_to", text="To")
        col.active = operator.do_translate


def frame_to_ms(frame):
    """Converts a frame number to a time in milliseconds, taking the frame rate of the scene into account."""
    scene = bpy.context.scene
    fps = (
        scene.render.fps / scene.render.fps_base
    )  # Calculate the frame rate as frames per second.
    ms_per_frame = 1000 / fps  # Calculate the number of milliseconds per frame.
    return frame * ms_per_frame


class SEQUENCER_OT_export_list_subtitles(Operator, ImportHelper):
    """Export Subtitles"""

    bl_idname = "sequencer.export_list_subtitles"
    bl_label = "Export"
    bl_options = {"REGISTER", "UNDO"}

    filename_ext = [".srt", ".ass"]

    filter_glob: StringProperty(
        default="*.srt;*.ass;*.ssa;*.mpl2;*.tmp;*.vtt;*.microdvd;*.fountain",
        options={"HIDDEN"},
        maxlen=255,
    )

    formats: EnumProperty(
        name="Formats",
        description="Format to save as",
        items=(
            ("srt", "srt", "SubRip"),
            ("ass", "ass", "SubStationAlpha"),
            ("ssa", "ssa", "SubStationAlpha"),
            ("mpl2", "mpl2", "MPL2"),
            # ("tmp", "tmp", "TMP"),
            ("vtt", "vtt", "WebVTT"),
            # ("microdvd", "microdvd", "MicroDVD"),
            ("fountain", "fountain", "Fountain Screenplay"),
        ),
        default="srt",
    )

    def execute(self, context):
        #ensure_pip(self)
        print("Please wait. Checking pysubs2 module...")
                
        if not import_module(self, "pysubs2", "pysubs2"):
            return {"CANCELLED"}
        import pysubs2
        # Get a list of all Text Strips in the VSE
        text_strips = [
            strip
            for strip in bpy.context.scene.sequence_editor.sequences
            if strip.type == "TEXT"
        ]

        # Sort the Subtitle Editor based on their start times in the timeline
        text_strips.sort(key=lambda strip: strip.frame_start)

        from pysubs2 import SSAFile, SSAEvent, make_time

        file_name = self.filepath
        if pathlib.Path(file_name).suffix != "." + self.formats:
            file_name = self.filepath + "." + self.formats
        if self.formats != "fountain":
            subs = SSAFile()
            # Iterate through the sorted text strips and add them to the list
            for strip in text_strips:
                event = SSAEvent()
                event.start = frame_to_ms(strip.frame_final_start)
                event.end = frame_to_ms(
                    strip.frame_final_start + strip.frame_final_duration
                )
                event.text = strip.text
                event.bold = strip.use_bold
                event.italic = strip.use_italic

                subs.append(event)
            text = subs.to_string(self.formats)
            #            if self.formats == "microdvd": #doesn't work
            #                subs.save(file_name, format_="microdvd", fps=(scene.render.fps / scene.render.fps_base))
            if self.formats == "mpl2":
                subs.save(file_name, format_="mpl2")
            else:
                subs.save(file_name)
        else:
            text = ""
            # Iterate through the sorted text strips and add them to the list
            for strip in text_strips:
                text = text + strip.text + chr(10) + chr(13) + " " + chr(10) + chr(13)
            fountain_file = open(file_name, "w")
            fountain_file.write(text)
            fountain_file.close()
        return {"FINISHED"}

    def draw(self, context):
        pass


class SEQUENCER_PT_export_list_subtitles(bpy.types.Panel):
    bl_space_type = "FILE_BROWSER"
    bl_region_type = "TOOL_PROPS"
    bl_label = "File Formats"
    bl_parent_id = "FILE_PT_operator"

    @classmethod
    def poll(cls, context):
        sfile = context.space_data
        operator = sfile.active_operator
        return operator.bl_idname == "SEQUENCER_OT_export_list_subtitles"

    def draw(self, context):
        layout = self.layout
        layout.use_property_split = True
        layout.use_property_decorate = False

        sfile = context.space_data
        operator = sfile.active_operator
        layout = layout.column(heading="File Formats")
        layout.prop(operator, "formats", text="Format")


class SEQUENCER_OT_copy_textprops_to_selected(Operator):
    """Copy properties from active text strip to selected text strips"""

    bl_idname = "sequencer.copy_textprops_to_selected"
    bl_label = "Copy Properties to Selected"
    bl_options = {"REGISTER", "UNDO"}

    def execute(self, context):
        current_scene = bpy.context.scene
        try:
            active = current_scene.sequence_editor.active_strip
        except AttributeError:
            self.report({"INFO"}, "No active strip selected")
            return {"CANCELLED"}
        for strip in context.selected_sequences:
            if strip.type == active.type == "TEXT":
                strip.wrap_width = active.wrap_width
                strip.font = active.font
                strip.use_italic = active.use_italic
                strip.use_bold = active.use_bold
                strip.font_size = active.font_size
                strip.color = active.color
                strip.use_shadow = active.use_shadow
                strip.shadow_color = active.shadow_color
                strip.use_box = active.use_box
                strip.box_color = active.box_color
                strip.box_margin = active.box_margin
                strip.location[0] = active.location[0]
                strip.location[1] = active.location[1]
                strip.anchor_x = active.anchor_x
                strip.anchor_y = active.anchor_y
        return {"FINISHED"}


# Define the panel to hold the UIList and the refresh button
class SEQUENCER_PT_panel(bpy.types.Panel):
    bl_idname = "SEQUENCER_PT_panel"
    bl_label = "Subtitle Editor"
    bl_space_type = "SEQUENCE_EDITOR"
    bl_region_type = "UI"
    bl_category = "Subtitle Editor"

    def draw(self, context):
        layout = self.layout

        row = layout.row()
        col = row.column()
        col.template_list(
            "SEQUENCER_UL_List",
            "",
            context.scene,
            "text_strip_items",
            context.scene,
            "text_strip_items_index",
            rows=14,
        )

        row = row.column(align=True)
        row.operator("text.refresh_list", text="", icon="FILE_REFRESH")

        row.separator()

        row.operator("sequencer.import_subtitles", text="", icon="IMPORT")
        row.operator("sequencer.export_list_subtitles", text="", icon="EXPORT")

        row.separator()

        row.operator("text.add_strip", text="", icon="ADD", emboss=True)
        row.operator("text.delete_item", text="", icon="REMOVE", emboss=True)
        row.operator("text.delete_strip", text="", icon="SCULPTMODE_HLT", emboss=True)

        row.separator()

        row.operator("text.select_previous", text="", icon="TRIA_UP")
        row.operator("text.select_next", text="", icon="TRIA_DOWN")

        row.separator()

        row.operator("text.insert_newline", text="", icon="EVENT_RETURN")


class SEQUENCER_PT_whisper_panel(bpy.types.Panel):
    """UI Panel in the Sequencer's Strip Properties"""
    bl_label = "Transcription & Translation"
    bl_idname = "SEQUENCER_PT_whisper"
    bl_space_type = 'SEQUENCE_EDITOR'
    bl_region_type = "UI"
    bl_category = "Subtitle Editor"

    @classmethod
    def poll(cls, context):
        # Only show the panel if a strip is selected (or always show?)
        # return context.selected_sequences is not None
        # Let's always show it for easier access to Setup
        return context.scene is not None


    def draw(self, context):
        layout = self.layout
        scene = context.scene
        props = scene.whisper_props # Get property group instance

        # --- Setup Section ---
        box = layout.box()
        row = box.row(align=True)
        # Display status icon based on check
        status_icon = 'CHECKMARK' if dependencies_installed else 'ERROR'
        if not dependencies_checked and not dependencies_installed:
             status_icon = 'QUESTION' # Indicate unchecked state

        row.label(text="Dependencies:", icon=status_icon)
        row.operator(SEQUENCER_OT_whisper_setup.bl_idname, icon='SCRIPTPLUGINS')

        # Disable transcription controls if dependencies not met
        is_ready = dependencies_installed and faster_whisper_module is not None
        col = layout.column()
        col.enabled = is_ready # Disable subsequent sections if not ready

        # --- Configuration Section ---
        box = col.box()
        # ... (existing model, device, compute, language, beam, vad props) ...
        box.prop(props, "model_size", text="Model")
        row = box.row(align=True)
        row.prop(props, "device", text="Device")
        row.prop(props, "compute_type", text="Compute")
        box.prop(props, "language", text="Language")
        row = box.row(align=True)
        row.prop(props, "beam_size", text="Beam Size")
        row.prop(props, "use_vad", text="VAD Filter")


        # --- NEW: Subtitle Output Settings ---
        box = col.box()
        row = box.row(align=True)
        row.prop(props, "output_channel", text="Channel")
        row.prop(props, "font_size", text="Font Size")
        row = box.row(align=True)
        row.prop(props, "text_align_y", text="")
        row.prop(props, "wrap_width", text="Wrap Width")


        # --- Actions Section ---
        box = col.box()
        action_col = box.column(align=True)

        strip_selected = get_selected_strip(context) is not None and SEQUENCER_OT_whisper_transcribe.poll(context)
        action_col.enabled = strip_selected

        # UPDATE BUTTON TEXTS
        op_transcribe = action_col.operator(SEQUENCER_OT_whisper_transcribe.bl_idname, text="Transcribe to Text Strips", icon='REC')
        op_transcribe.task = "transcribe"

        op_translate = action_col.operator(SEQUENCER_OT_whisper_transcribe.bl_idname, text="Translate to Text Strips (EN)", icon='WORDWRAP_ON')
        op_translate.task = "translate"


def import_subtitles(self, context):
    layout = self.layout
    layout.separator()
    layout.operator("sequencer.import_subtitles", text="Subtitles", icon="ALIGN_BOTTOM")


def transcribe(self, context):
    layout = self.layout
    layout.separator()
    layout.operator("text.transcribe", text="Transcriptions", icon="SPEAKER")


def copyto_panel_append(self, context):
    strip = context.active_sequence_strip
    strip_type = strip.type
    if strip_type == "TEXT":
        layout = self.layout
        layout.operator(SEQUENCER_OT_copy_textprops_to_selected.bl_idname)


def setText(self, context):
    scene = context.scene
    current_index = context.scene.text_strip_items_index
    max_index = len(context.scene.text_strip_items)

    if current_index < max_index:
        selected_item = scene.text_strip_items[scene.text_strip_items_index]
        selected_item.select = True
        update_text(selected_item, context)


classes = (
    SEQUENCER_PT_panel,
    subtitle_preferences,
    TextStripItem,
    SEQUENCER_UL_List,
    SEQUENCER_OT_refresh_list,
    SEQUENCER_OT_add_strip,
    SEQUENCER_OT_delete_item,
    SEQUENCER_OT_delete_strip,
    SEQUENCER_OT_select_next,
    SEQUENCER_OT_select_previous,
    SEQUENCER_OT_insert_newline,
    SEQUENCER_OT_import_subtitles,
    SEQUENCER_PT_import_subtitles,
    SEQUENCER_OT_export_list_subtitles,
    SEQUENCER_PT_export_list_subtitles,
    SEQUENCER_OT_copy_textprops_to_selected,
    TEXT_OT_transcribe,
    WhisperProperties,
    SEQUENCER_OT_whisper_setup,
    SEQUENCER_OT_whisper_transcribe,
    SEQUENCER_PT_whisper_panel,
)


def register():
    print(f"Registering {bl_info['name']} Addon")
    # Add the property group to the Scene type
    for cls in classes:
        try:
            bpy.utils.register_class(cls)
        except ValueError as e:
             print(f"Class {cls.__name__} already registered? Skipping. Error: {e}")
        except Exception as e:
             print(f"Failed to register class {cls.__name__}: {e}")
    bpy.types.Scene.whisper_props = PointerProperty(type=WhisperProperties)

    # Reset global flags on registration / Blender start
    global dependencies_checked, dependencies_installed, faster_whisper_module
    dependencies_checked = False
    dependencies_installed = False
    faster_whisper_module = None
    # Perform an initial check silently on registration/startup
    print("Performing initial dependency check on startup...")
    check_faster_whisper()
    dependencies_checked = True # Mark as checked
    
    bpy.types.Scene.text_strip_items = bpy.props.CollectionProperty(type=TextStripItem)
    bpy.types.Scene.text_strip_items_index = bpy.props.IntProperty(
        name="Index for Subtitle Editor", default=0, update=setText
    )
    bpy.types.SEQUENCER_MT_add.append(import_subtitles)
    #bpy.types.SEQUENCER_MT_add.append(transcribe)
    bpy.types.SEQUENCER_PT_effect.append(copyto_panel_append)


def unregister():
    bpy.types.SEQUENCER_PT_effect.remove(copyto_panel_append)
    print(f"Unregistering {bl_info['name']} Addon")
    # Delete the property group from Scene first
    del Scene.whisper_props

    for cls in reversed(classes):
        try:
            bpy.utils.unregister_class(cls)
        except RuntimeError as e:
            print(f"Failed to unregister class {cls.__name__}: {e}")
        except Exception as e:
             print(f"Error unregistering class {cls.__name__}: {e}")

    # Clear globals on unregister (optional, good practice)
    global dependencies_checked, dependencies_installed, faster_whisper_module
    dependencies_checked = False
    dependencies_installed = False
    faster_whisper_module = None
    del bpy.types.Scene.text_strip_items
    del bpy.types.Scene.text_strip_items_index
    bpy.types.SEQUENCER_MT_add.remove(import_subtitles)
    #bpy.types.SEQUENCER_MT_add.append(transcribe)


# Register the addon when this script is run
if __name__ == "__main__":
    register()
