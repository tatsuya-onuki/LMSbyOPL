import datetime
import threading
import tkinter as tk
from tkinter import messagebox

import gspread
import nfc
from oauth2client.service_account import ServiceAccountCredentials


class NFCReader:
    SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1BsKymbniKnSCFJAkYbflLoAjd4GzjSRC17NU-hhkFlo/edit#gid=0"
    JSON_KEYFILE_PATH = R"C:\Users\oonuk\OneDrive\デスクトップ\python_JLD\lms-dx-b724fbd23653.json"

    def __init__(self, reader_identifier):
        self.reader_identifier = reader_identifier
        self.last_connect_times = {}
        self.tag_id = None
        self.dt_now = None
        self.is_running = True

        try:
            self.clf = nfc.ContactlessFrontend('usb')
        except IOError:
            self.clf = None
            print("NFC Reader not connected")

        self.root = tk.Tk()

        # Define the size and position of the root window
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width = int(screen_width * 0.8)  # 80% of the screen width
        height = int(screen_height * 0.8)  # 80% of the screen height
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f'{width}x{height}+{x}+{y}')

        self.root.configure(bg="#0054A6")
        self.root.title("Process Selection")

        self.create_start_frame()

        if self.clf is not None:
            self.nfc_thread = threading.Thread(target=self.start_nfc_loop, daemon=True)
            self.nfc_thread.start()

        self.barcode_window = None
        self.barcode_entry = None

        self.root.mainloop()

        if self.clf is not None:
            self.clf.close()


    def get_worksheet(self, sheet_name):
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(self.JSON_KEYFILE_PATH, scope)
        gc = gspread.authorize(credentials)
        return gc.open_by_url(self.SPREADSHEET_URL).worksheet(sheet_name)

    def create_custom_button(self, *args, **kwargs):
        command = kwargs.pop('command', None)
        bg = kwargs.get('bg', '#1A4472')
        hover_bg = '#666666'

        def on_enter(e):
            button['cursor'] = 'hand2'
            button['bg'] = hover_bg

        def on_leave(e):
            button['cursor'] = ''
            button['bg'] = bg

        button = tk.Button(*args, **kwargs, command=command)
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

        return button

    def create_start_frame(self):
        self.start_frame = tk.Frame(self.root, bg="#0054A6", width=1200, height=400)
        self.start_frame.place(relx=0.5, rely=0.5, anchor='center')

        start_label = tk.Label(self.start_frame, text="IDカードをリーダーにタッチしてください", bg="#0054A6", fg="white", font=("Helvetica", 30))
        start_label.grid(row=0, column=0, columnspan=5, pady=40)  # Adding pady here

        no_id_label = tk.Label(self.start_frame, text="IDがない方はこちらからお選びください", bg="#0054A6", fg="white", font=("Helvetica", 14))
        no_id_label.grid(row=1, column=0, columnspan=5, pady=20)  # Adding pady here

        for i in range(10):
            no_id_button = self.create_custom_button(self.start_frame, text=f"{i+1}", command=lambda i=i: self.switch_to_id_frame(i+1), fg="white", width=8, height=2, font=("Helvetica", 22), bg="#1A4472")
            no_id_button.grid(row=2+i//5, column=i%5, padx=2, pady=6)  # Reducing padx and pady here


    def create_id_frame(self, id=None):
        self.id_frame = tk.Frame(self.root, bg="#0054A6", width=1200, height=400)
        self.id_frame.place(relx=0.5, rely=0.1, anchor='center')

        self.id_label = tk.Label(self.id_frame, text="ID: ", bg="#0054A6", fg="white", font=("Helvetica", 24))
        self.id_label.pack(expand=True, fill='both', padx=40, pady=40)

        if id is not None:
            self.id_label.config(text=f"ID: {id}")

    def switch_to_id_frame(self, id_number):
        self.start_frame.destroy()
        self.tag_id = f"{id_number}"
        self.dt_now = datetime.datetime.now()
        self.create_id_frame(id=self.tag_id)
        self.create_button_frame()

    def create_button_frame(self):
        self.button_frame = tk.Frame(self.root, bg="#0054A6")
        self.button_frame.place(relx=0.5, rely=0.5, anchor='center')

        processes = ['入庫検品', 'ぷちぷち加工', '封入作業', '格納', 'TotalPick', 'SinglePick', '出庫検品', '梱包', '出庫', '休憩', '終了']
        for i, process_name in enumerate(processes):
            if process_name == 'TotalPick' or process_name == 'SinglePick':
                button = self.create_custom_button(self.button_frame, text=process_name, command=lambda name=process_name: self.show_barcode_entry(name), fg="white", width=30, height=5, font=("Helvetica", 18), bg="#1A4472")
                button.grid(row=i // 3, column=i % 3, padx=5, pady=5)
            else:
                button = self.create_custom_button(self.button_frame, text=process_name, command=lambda name=process_name: (self.record_process(name), self.switch_to_start_frame()), fg="white", width=30, height=5, font=("Helvetica", 18), bg="#1A4472")
                button.grid(row=i // 3, column=i % 3, padx=5, pady=5)

    def switch_to_start_frame(self):
        self.id_frame.destroy()
        self.button_frame.destroy()
        self.create_start_frame()

    def show_barcode_entry(self, process_name):
        if process_name == 'TotalPick' or process_name == 'SinglePick':
            self.barcode_window = tk.Toplevel(self.root)
            self.barcode_window.attributes('-fullscreen', True)
            self.barcode_window.configure(bg="#0054A6")
            self.barcode_window.title("Barcode Entry")

            self.create_barcode_entry_frame(process_name)

    def create_barcode_entry_frame(self, process_name):
        self.barcode_entry_frame = tk.Frame(self.barcode_window, bg="#0054A6", width=1200, height=400)
        self.barcode_entry_frame.place(relx=0.5, rely=0.5, anchor='center')

        barcode_label = tk.Label(self.barcode_entry_frame, text="Work IDバーコードをスキャンしてください", bg="#0054A6", fg="white", font=("Helvetica", 24))
        barcode_label.pack(expand=True, fill='both', padx=40, pady=40)

        self.barcode_entry = tk.Entry(self.barcode_entry_frame, font=("Helvetica", 24))
        self.barcode_entry.pack(padx=10, pady=10)
        self.barcode_entry.bind('<Return>', lambda event: self.submit_barcode(process_name))  # Enter key is bound to the submit_barcode function

        submit_button = tk.Button(self.barcode_entry_frame, text="Submit", command=lambda: self.submit_barcode(process_name), fg="white", font=("Helvetica", 18), bg="#1A4472")
        submit_button.pack(padx=10, pady=10)


    def submit_barcode(self, process_name):
        work_id = self.barcode_entry.get()
        self.barcode_entry.delete(0, tk.END)
        self.barcode_window.destroy()

        self.dt_now = datetime.datetime.now()
        self.record_process(process_name, work_id=work_id)

        self.switch_to_start_frame()

    def on_connect(self, tag):
        self.dt_now = datetime.datetime.now()
        self.tag_id = tag.identifier.hex()

        if self.tag_id in self.last_connect_times and (self.dt_now - self.last_connect_times[self.tag_id]).total_seconds() < 10:
            print("Same tag connected too soon.")
            return

        self.last_connect_times[self.tag_id] = self.dt_now

        print("Tag connected: ", self.tag_id)

        self.root.attributes('-topmost', True)  # Add this line 

        self.start_frame.destroy()
        self.create_id_frame()
        self.create_button_frame()

        self.id_label.config(text=f"ID: {self.tag_id}")

        self.root.attributes('-topmost', False)  # Add this line


    def record_process(self, process_name, work_id=None):
        if self.tag_id and self.dt_now:
            if process_name == '終了':
                self.calculate_time_difference()
                print(f"Recording process {process_name} for tag {self.tag_id} at {self.dt_now}")
                self.write_spread_sheet("sheet1", [self.reader_identifier, self.tag_id, self.dt_now.strftime('%Y-%m-%d %H:%M:%S'), process_name, work_id])
                self.tag_id = None
                self.dt_now = None

                self.id_label.config(text="ID: ")
            else:
                print(f"Recording process {process_name} for tag {self.tag_id} at {self.dt_now}")
                self.write_spread_sheet("sheet1", [self.reader_identifier, self.tag_id, self.dt_now.strftime('%Y-%m-%d %H:%M:%S'), process_name, work_id])

    def write_spread_sheet(self, sheet_name, str_contents):
        worksheet = self.get_worksheet(sheet_name)
        worksheet.append_row(str_contents)

    def calculate_time_difference(self):
        worksheet1 = self.get_worksheet("sheet1")
        data = worksheet1.get_all_values()

        last_entry = next((entry for entry in reversed(data) if entry[1] == self.tag_id), None)
        if last_entry is None:
            print(f"作業不明 {self.tag_id}")
            return

        start_time = datetime.datetime.strptime(last_entry[2], '%Y-%m-%d %H:%M:%S')
        end_time = self.dt_now
        seconds = int((end_time - start_time).total_seconds())
        minutes, seconds = divmod(seconds, 60)
        hours, minutes = divmod(minutes, 60)
        diff = datetime.time(hour=hours, minute=minutes, second=seconds)



        # worksheet2 = self.get_worksheet("sheet2")
        # worksheet2.append_row([self.reader_identifier, self.tag_id, last_entry[3], start_time.strftime('%Y-%m-%d %H:%M:%S'), end_time.strftime('%Y-%m-%d %H:%M:%S'), str(diff)])
        worksheet2 = self.get_worksheet("sheet2")
        worksheet2.append_row([self.reader_identifier, self.tag_id, last_entry[3], start_time.strftime('%Y-%m-%d %H:%M:%S'), end_time.strftime('%Y-%m-%d %H:%M:%S'), str(diff), last_entry[4]])

    def start_nfc_loop(self):
        if self.clf is not None:
            try:
                while self.is_running:
                    tag = self.clf.connect(rdwr={'on-connect': self.on_connect, 'targets': ['106A']}, terminate=lambda: not self.is_running)
                    if tag:
                        print("Tag connected")
            finally:
                self.is_running = False


    def exit_program(self):
        if messagebox.askokcancel("Quit", "本当にLMSを終了しますか?"):
            self.root.destroy()

    def create_close_button(self):
        close_button = tk.Button(self.root, text="X", command=self.exit_program, bg="#0054A6", fg="white", font=("Helvetica", 12), borderwidth=0, height=1, width=2)
        close_button.place(relx=0.975, rely=0.02, anchor='center')


if __name__ == "__main__":
    reader = NFCReader('端末1')
    