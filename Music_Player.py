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

        Prev_Track_Btn = tk.Button(self.parent, text="Previous Track", command=self.Prev_Track)
        Prev_Track_Btn.place(relx=0.3, rely=0.5, anchor='center')

        Next_P_Btn = tk.Button(self.parent, text="Next Playlist", command=self.Next_Playlist)
        Next_P_Btn.place(relx=0.7, rely=0.8, anchor='center')

        Prev_P_Btn = tk.Button(self.parent, text="Previous Playlist", command=self.Prev_Playlist)
        Prev_P_Btn.place(relx=0.3, rely=0.8, anchor='center')

        pass

    def Next_Playlist(self):
        if Playlist.playNum >= (Playlist.WSNum - 1):
            Playlist.playNum = 0
            pass
        else:
            Playlist.playNum += 1

        self.ActivePlaylist = Playlist.wb.sheetnames[Playlist.playNum]

        Playlist.FirstEmpty = None
        num = 2
        while Playlist.FirstEmpty is None:
            if Playlist.wb[Playlist.wb.sheetnames[Playlist.playNum]]['B' + str(num)].value is None:
                Playlist.FirstEmpty = 'B' + str(num)
                pass
            else:
                num += 1
                pass

        pass

    def Prev_Playlist(self):
        if Playlist.playNum >= (Playlist.WSNum - 1):
            Playlist.playNum = 0
            pass
        else:
            Playlist.playNum += 1

        self.ActivePlaylist = Playlist.wb.sheetnames[Playlist.playNum]

        Playlist.FirstEmpty = None
        num = 2
        while Playlist.FirstEmpty is None:
            if Playlist.wb[Playlist.wb.sheetnames[Playlist.playNum]]['B' + str(num)].value is None:
                Playlist.FirstEmpty = 'B' + str(num)
                pass
            else:
                num += 1
                pass
        
        pass

    def Next_Track(self):
        Playlist.trackNum += 1
        if ('B' + str(Playlist.trackNum)) == Playlist.FirstEmpty:
            Playlist.trackNum = 2
            pass
        else:
            pass

    def Prev_Track(self):
        Playlist.trackNum += 1
        if ('B' + str(Playlist.trackNum)) == Playlist.FirstEmpty:
            Playlist.trackNum = 2
            pass
        else:
            pass

    
class Playlist():
    def __init__(self):
        self.filepath = os.path.dirname(os.path.abspath(__file__)) + "\Music\Playlists.xlsx"
        self.playNum = 0
        self.trackNum = 2

        self.wb = openpyxl.load_workbook(self.filepath,True)
        
        self.WSNum = len(self.wb.sheetnames)

        self.FirstEmpty = None
        num = 2
        while self.FirstEmpty is None:
            if self.wb[self.wb.sheetnames[self.playNum]]['B' + str(num)].value is None:
                self.FirstEmpty = 'B' + str(num)
                pass
            else:
                num += 1
                pass

        pass


if __name__ == "__main__":
    Playlist = Playlist()

    root = tk.Tk()
    MainApp = MainApplication(root)
    MainApp.pack(side="top", fill="both", expand=False)
    root.mainloop()