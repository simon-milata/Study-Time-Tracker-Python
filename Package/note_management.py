import customtkinter as ctk

from .styles import *

class NotesManager:
    def __init__(self, app, data_manager):
        self.app = app
        self.data_manager = data_manager

    def create_task(self, index):
        self.frame = ctk.CTkFrame(self.app.notes_data_frame, width=WIDTH, fg_color=(light_frame_color, frame_color), height=button_height + frame_padding * 2)
        self.frame.pack(pady=frame_padding/2)
        self.frame.grid_propagate(False)

        def title_text(title: str) -> str:
            max_letters = 40

            for letter in title:
                if letter.isupper():
                    max_letters -= 0.35

            max_letters = int(max_letters)

            display_title = title

            if len(title) > max_letters:
                display_title = title[:max_letters] + "..."
                
            return display_title

        title = ctk.CTkLabel(self.frame, text=title_text(str(self.data_manager.worksheet["O" + str(index)].value)), font=(font_family, font_size*1.2),
                             text_color=(light_font_color, font_color), anchor="center", height=button_height + frame_padding * 2)
        title.grid(row=0, column=0, padx=widget_padding_x)

        date = ctk.CTkLabel(self.frame, text=str(self.data_manager.worksheet["N" + str(index)].value), font=(font_family, font_size*1),
                            text_color=(light_off_font_color, off_font_color), anchor="center", height=button_height + frame_padding * 2)
        date.place(anchor="center", relx=0.575, rely=0.5)

        button_frame = ctk.CTkFrame(self.frame, fg_color="transparent")
        button_frame.place(anchor="center", rely=0.5, relx=0.825)

        open_button = ctk.CTkButton(button_frame, text="Open", height=button_height, fg_color=self.data_manager.color, hover_color=self.data_manager.highlight_color, font=(font_family, font_size), text_color=button_font_color, 
                                   command=lambda: self._open_notes_text(str(self.data_manager.worksheet["N" + str(index)].value), str(self.data_manager.worksheet["O" + str(index)].value), str(self.data_manager.worksheet["P" + str(index)].value), index)
                                   ,anchor="center")
        open_button.grid(row=0, column=0, padx=widget_padding_x)
        delete_button = ctk.CTkButton(button_frame, text="Delete", height=button_height, fg_color=self.data_manager.color, hover_color=self.data_manager.highlight_color, font=(font_family, font_size), text_color=button_font_color,
                                          command=lambda: self.delete_task(index), anchor="center")
        delete_button.grid(row=0, column=1)


    def delete_task(self, index):
        self.data_manager.worksheet["M" + str(index)].value = "Yes"
        self.frame.destroy()
        self.data_manager.load_notes()


    def _open_notes_text(self, date, title, text, index):
        self.app.notes_frame_frame.grid_forget()

        frame = ctk.CTkFrame(self.app.notes_frame, fg_color="transparent", corner_radius=10, height=HEIGHT + frame_padding * 2, width=WIDTH - frame_padding * 2)
        frame.grid(padx=frame_padding, pady=frame_padding)
        frame.grid_propagate(False)

        header_frame = ctk.CTkFrame(frame, fg_color=(light_frame_color, frame_color), corner_radius=10, width=WIDTH - frame_padding * 2, height=60)
        header_frame.grid(row=0, column=0, pady=(0, frame_padding))
        header_frame.grid_propagate(False)

        def title_text(title: str) -> str:
            max_letters = 45

            for letter in title:
                if letter.isupper():
                    max_letters -= 0.4
            

            max_letters = int(max_letters)

            display_title = title

            if len(title) > max_letters:
                display_title = title[:max_letters] + "..."
                
            return display_title

        title_label = ctk.CTkLabel(header_frame, text=title_text(title), font=(font_family, font_size*1.5), text_color=(light_font_color, font_color),
                                   height=40, width=WIDTH - 280 - frame_padding * 6, fg_color=(light_frame_color, frame_color), anchor="w")
        title_label.grid(row=0, column=0, padx=widget_padding_x, pady=widget_padding_y)

        button_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        button_frame.place(anchor="center", rely=0.5, relx=0.825)

        textbox = ctk.CTkTextbox(frame, font=(font_family, font_size), text_color=(light_font_color, font_color), fg_color="transparent", width=WIDTH - frame_padding * 2, height=HEIGHT + frame_padding * 2 - 60)
        textbox.grid(row=1, column=0)
        textbox.insert("0.0", text)
        textbox.configure(state="disabled")
        
        edit_button = ctk.CTkButton(button_frame, height=button_height, text="Edit", fg_color=self.data_manager.color, anchor="center",
                                                hover_color=self.data_manager.highlight_color, font=(font_family, font_size), text_color=button_font_color, command=lambda: self.edit_note(index, textbox, edit_button, frame))
        edit_button.grid(row=0, column=1)
        exit_button = ctk.CTkButton(button_frame, height=button_height, text="Exit", fg_color=self.data_manager.color, command=lambda: self._exit_note(frame), 
                                                hover_color=self.data_manager.highlight_color, font=(font_family, font_size), text_color=button_font_color, anchor="center")
        exit_button.grid(row=0, column=2, padx=widget_padding_x, pady=widget_padding_y)


    def _exit_note(self, frame):
        frame.destroy()
        self.app.notes_frame_frame.grid(row=0, column=0, padx=frame_padding, pady=(frame_padding, 0))


    def edit_note(self, index, textbox, edit_button, frame):
        def save_note(index, frame):
            self.data_manager.worksheet["P" + str(index)].value = textbox.get("0.0", "end")
            self._exit_note(frame)
            self.data_manager.workbook.save(self.app.data_file)

        textbox.configure(state="normal", fg_color=(light_frame_color, frame_color))
        edit_button.configure(command=lambda: save_note(index, frame), text="Save")


    def clear_notes(self):
        for note in self.app.notes_data_frame.winfo_children():
            note.destroy()