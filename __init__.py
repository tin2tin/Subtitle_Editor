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
    "version": (1, 0),
    "blender": (2, 30, 0),
    "location": "Sequencer > Side Bar > Subtitle Editor Tab",
    "warning": "",
    "doc_url": "",
    "support": "COMMUNITY",
    "category": "Sequencer",
}

import os, sys, bpy, pathlib, re
from bpy.types import Operator
from bpy_extras.io_utils import ImportHelper
from bpy.props import StringProperty, BoolProperty, EnumProperty


def get_strip_by_name(name):
    for strip in bpy.context.scene.sequence_editor.sequences[0:]:
        if strip.name == name:
            print(strip.text)
            return strip
    return None  # Return None if the strip is not found


import bpy

def find_first_empty_channel(start_frame, end_frame):
    for ch in range(1, len(bpy.context.scene.sequence_editor.sequences_all) + 1):
        for seq in bpy.context.scene.sequence_editor.sequences_all:
            if seq.channel == ch and seq.frame_final_start < end_frame and seq.frame_final_end > start_frame:
                break
        else:
            return ch
    return 1


def update_text(self, context):
    for strip in bpy.context.scene.sequence_editor.sequences[0:]:
        if strip.type == "TEXT" and strip.name == self.name:
            # Update the text of the text strip
            strip.text = self.text

            break
        strip = get_strip_by_name(self.name)
        if strip:
            # Deselect all strips.
            for seq in context.scene.sequence_editor.sequences_all:
                seq.select = False
            bpy.context.scene.sequence_editor.active_strip = strip
            strip.select = True

            # Set the current frame to the start frame of the active strip
            bpy.context.scene.frame_set(int(strip.frame_start))


# Define a custom property group to hold the text strip name and text
class TextStripItem(bpy.types.PropertyGroup):
    name: bpy.props.StringProperty()
    text: bpy.props.StringProperty(update=update_text)
    selected: bpy.props.IntProperty()


class SEQUENCER_UL_List(bpy.types.UIList):
    def draw_item(
        self, context, layout, data, item, icon, active_data, active_propname
    ):
        layout.prop(item, "text", text="", emboss=False)

    def invoke(self, context, event):  # doesn't work
        if event.type == "LEFTMOUSE" and event.value == "RELEASE":
            context.scene.text_strip_items_index = item.index = self.layout.active_index
            selected_item = context.scene.text_strip_items[
                context.scene.text_strip_items_index
            ]
            update_text(selected_item, context)
        return {"RUNNING_MODAL"}


# Define an operator to refresh the list
class SEQUENCER_OT_refresh_list(bpy.types.Operator):
    """Sync items in the list with the text strips"""
    bl_idname = "text.refresh_list"
    bl_label = "Refresh List"

    def execute(self, context):
        active = context.scene.sequence_editor.active_strip
        # Clear the list
        context.scene.text_strip_items.clear()

        # Get a list of all Subtitle Editor in the VSE
        text_strips = [
            strip
            for strip in bpy.context.scene.sequence_editor.sequences
            if strip.type == "TEXT"
        ]

        # Sort the Subtitle Editor based on their start times in the timeline
        text_strips.sort(key=lambda strip: strip.frame_start)

        # Iterate through the sorted Subtitle Editor and add them to the list
        for strip in text_strips:
            item = context.scene.text_strip_items.add()
            item.name = strip.name
            item.text = strip.text
            # context.scene.text_strip_items_index +=1
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
            out_frame = cf + (2*strip.frame_final_duration)
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
                new_strip.box_margin = strip.box_margin
                new_strip.location = strip.location
                new_strip.align_x = strip.align_x
                new_strip.align_y = strip.align_y
                context.scene.sequence_editor.active_strip = new_strip
        else:
            strips = scene.sequence_editor.sequences
            chan = find_first_empty_channel(context.scene.frame_current, context.scene.frame_current+100)

            new_strip = strips.new_effect(
                name="Text",
                type="TEXT",
                channel=chan,
                frame_start=context.scene.frame_current,
                frame_end=context.scene.frame_current+100,
            )
            context.scene.sequence_editor.active_strip = new_strip
            new_strip.select = True
        # Refresh the UIList
        bpy.ops.text.refresh_list()

#        scene.text_strip_items_index = len(scene.text_strip_items)
#        selected_item = scene.text_strip_items[scene.text_strip_items_index]#scene.text_strip_items_index]
#        selected_item.select = True
#        context.scene.text_strip_items_index = len(context.scene.text_strip_items) - 1

        # Select the new item in the UIList
        context.scene.text_strip_items_index = index + 1

        return {"FINISHED"}


class SEQUENCER_OT_delete_strip(bpy.types.Operator):
    """Remove item and strip"""
    bl_idname = "text.delete_strip"
    bl_label = "Remove Item & Strip"

    @classmethod
    def poll(cls, context):
        index = context.scene.text_strip_items_index
        items = context.scene.text_strip_items
        selected = len(items) - 1 < index
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
        return {"FINISHED"}


class SEQUENCER_OT_import_subtitles(Operator, ImportHelper):
    """Import Subtitles"""

    bl_idname = "sequencer.import_subtitles"
    bl_label = "Import"
    bl_options = {"REGISTER", "UNDO"}

    filename_ext = [".srt", ".ass"]

    filter_glob: StringProperty(
        default="*.srt;*.ass;*.ssa;*.mpl2;*.tmp;*.vtt;*.microdvd",
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
            print(self.translate_from)
            try:
                from srtranslator import SrtFile
                from srtranslator.translators.deepl import DeeplTranslator
            except ModuleNotFoundError:
                import site
                import subprocess
                import sys

                app_path = site.USER_SITE
                if app_path not in sys.path:
                    sys.path.append(app_path)
                pybin = sys.executable  # bpy.app.binary_path_python # Use for 2.83

                print("Ensuring: pip")
                try:
                    subprocess.call([pybin, "-m", "ensurepip"])
                except ImportError:
                    pass
                self.report({"INFO"}, "Installing: srtranslator module.")
                print("Installing: srtranslator module")
                subprocess.check_call([pybin, "-m", "pip", "install", "srtranslator"])
                try:
                    from srtranslator import SrtFile
                    from srtranslator.translators.deepl import DeeplTranslator

                    self.report({"INFO"}, "Detected: srtranslator module.")
                    print("Detected: srtranslator module")
                except ModuleNotFoundError:
                    print("Installation of the srtranslator module failed")
                    self.report(
                        {"INFO"},
                        "Installing srtranslator module failed! Try to run Blender as administrator.",
                    )
                    return {"CANCELLED"}
            file = self.filepath
            if not file:
                return {"CANCELLED"}
                print("Invalid file")
                self.report({"INFO"}, "Invalid file")
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
        try:
            import pysubs2
        except ModuleNotFoundError:
            import site
            import subprocess
            import sys

            app_path = site.USER_SITE
            if app_path not in sys.path:
                sys.path.append(app_path)
            pybin = sys.executable  # bpy.app.binary_path_python # Use for 2.83

            print("Ensuring: pip")
            try:
                subprocess.call([pybin, "-m", "ensurepip"])
            except ImportError:
                pass
            self.report({"INFO"}, "Installing: pysubs2 module.")
            print("Installing: pysubs2 module")
            subprocess.check_call([pybin, "-m", "pip", "install", "pysubs2"])
            try:
                import pysubs2
            except ModuleNotFoundError:
                print("Installation of the pysubs2 module failed")
                self.report(
                    {"INFO"},
                    "Installing pysubs2 module failed! Try to run Blender as administrator.",
                )
                return {"CANCELLED"}
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
        file = self.filepath
        if self.do_translate:
            file = f"{os.path.splitext(file)[0]}_translated.srt"
        if not file:
            return {"CANCELLED"}
        if pathlib.Path(file).is_file():
            if (
                pathlib.Path(file).suffix
                not in pysubs2.formats.FILE_EXTENSION_TO_FORMAT_IDENTIFIER
            ):
                print("Unable to extract subtitles from file")
                self.report({"INFO"}, "Unable to extract subtitles from file")
                return {"CANCELLED"}
        try:
            subs = pysubs2.load(file, fps=fps, encoding="utf-8")
        except:
            print("Import failed. Text encoding must be in UTF-8.")
            self.report({"INFO"}, "Import failed. Text encoding must be in UTF-8.")
            return {"CANCELLED"}
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
                print(pos_trim)
                pos_trim = pos_trim.group(1)
                print(pos_trim)
                pos = pos_trim.split(",")
                x = int(pos[0]) / render.resolution_x
                y = (render.resolution_y - int(pos[1])) / render.resolution_y
                position = True
                line.text = re.sub(r"{.+?}", "", line.text)
            print(line.start * fps_conv)
            new_strip = editor.sequences.new_effect(
                name=line.text,
                type="TEXT",
                channel=addSceneChannel,
                frame_start=int(line.start * fps_conv),
                frame_end=int(line.end * fps_conv),
            )
            new_strip.text = line.text
            new_strip.wrap_width = 0.68
            new_strip.font_size = 44
            new_strip.location[1] = 0.25
            new_strip.align_x = "CENTER"
            new_strip.align_y = "TOP"
            new_strip.use_shadow = True
            new_strip.use_box = True
            if position:
                new_strip.location[0] = x
                new_strip.location[1] = y
                new_strip.align_x = "LEFT"
            if italics:
                new_strip.use_italic = True
            if bold:
                new_strip.use_bold = True
        # Refresh the UIList
        bpy.ops.text.refresh_list()

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
                strip.align_x = active.align_x
                strip.align_y = active.align_y
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
            rows=1,
        )

        row = row.column(align=True)
        row.operator("text.refresh_list", text="", icon="FILE_REFRESH")

        row.separator()

        row.operator("sequencer.import_subtitles", text="", icon="IMPORT")
        row.operator("sequencer.export_subtitles", text="", icon="EXPORT")

        row.separator()

        row.operator("text.add_strip", text="", icon="ADD", emboss=True)
        row.operator("text.delete_strip", text="", icon="REMOVE", emboss=True)

        row.separator()

        row.operator("text.select_previous", text="", icon="TRIA_UP")
        row.operator("text.select_next", text="", icon="TRIA_DOWN")

        row.separator()

        row.operator("text.insert_newline", text="", icon="EVENT_RETURN")


def import_subtitles(self, context):
    layout = self.layout
    layout.separator()
    layout.operator("sequencer.import_subtitles", text="Subtitles", icon="ALIGN_BOTTOM")


def copyto_panel_append(self, context):
    strip = context.active_sequence_strip
    strip_type = strip.type
    if strip_type == "TEXT":
        layout = self.layout
        layout.operator(SEQUENCER_OT_copy_textprops_to_selected.bl_idname)


classes = (
    TextStripItem,
    SEQUENCER_UL_List,
    SEQUENCER_OT_refresh_list,
    SEQUENCER_OT_add_strip,
    SEQUENCER_OT_delete_strip,
    SEQUENCER_OT_select_next,
    SEQUENCER_OT_select_previous,
    SEQUENCER_OT_insert_newline,
    SEQUENCER_OT_import_subtitles,
    SEQUENCER_PT_import_subtitles,
    SEQUENCER_OT_copy_textprops_to_selected,
    SEQUENCER_PT_panel,
)


def register():
    for cls in classes:
        bpy.utils.register_class(cls)
    bpy.types.Scene.text_strip_items = bpy.props.CollectionProperty(type=TextStripItem)
    bpy.types.Scene.text_strip_items_index = bpy.props.IntProperty()
    bpy.types.SEQUENCER_MT_add.append(import_subtitles)
    bpy.types.SEQUENCER_PT_effect.append(copyto_panel_append)


def unregister():
    bpy.types.SEQUENCER_PT_effect.remove(copyto_panel_append)
    for cls in reversed(classes):
        bpy.utils.unregister_class(cls)
    del bpy.types.Scene.text_strip_items
    del bpy.types.Scene.text_strip_items_index
    bpy.types.SEQUENCER_MT_add.remove(import_subtitles)


# Register the addon when this script is run
if __name__ == "__main__":
    register()
