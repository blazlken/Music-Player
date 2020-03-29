import tkinter as tk
from pygame import mixer_music
import threading
import os
import openpyxl


class MainApplication(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent

        self.parent.title("Music Player")
        self.parent.geometry("300x500")

        Next_Track_Btn = tk.Button(self.parent, text="Next Track", command=self.Next_Track)
        Next_Track_Btn.place(relx=0.7, rely=0.5, anchor='center')

'''        Prev_Track_Btn = tk.Button(self.parent, text="Previous Track", command=self.Prev_Track)
        Prev_Track_Btn.place(relx=0.3, rely=0.5, anchor='center')

        Next_P_Btn = tk.Button(self.parent, text="Next Playlist", command=self.Next_Playlist)
        Next_P_Btn.place(relx=0.7, rely=0.8, anchor='center')

        Prev_P_Btn = tk.Button(self.parent, text="Previous Playlist", command=self.Prev_Playlist)
        Prev_P_Btn.place(relx=0.3, rely=0.8, anchor='center')
'''
        pass

    def Next_Track():
        trackNum += 1
        if Playlist().ActivePlaylist[]


class Playlist():
    def __init__(self):
        self.filepath = os.path.dirname(os.path.abspath(__file__)) + "\Music\Playlist.xlsx"

        wb = openpyxl.load_workbook(self.filepath,True)
        
        self.WSNum = len(wb.sheetnames)
        self.ActivePlaylist = wb[wb.sheetnames[self.WSNum]]

        pass

class Tracks():
    def __init__(self):

        self.FirstEmpty = None
        num = 2
        while self.FirstEmpty is None:
            if Playlist.ActivePlaylist['B' + str(num)] is None:
                self.FirstEmpty = 'B' + str(num)
                pass
            else:
                num += 1
                pass

if __name__ == "__main__":
    root = tk.Tk()
    MainApplication(root).pack(side="top", fill="both", expand=False)
    root.mainloop()