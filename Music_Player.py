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



class Playlist():
    def __init__(self):
        self.filepath = os.path.dirname(os.path.abspath(__file__)) + "\Music\Playlist.xlsx"
        self.wb = openpyxl.load_workbook(filename=self.filepath, read_only=True)
        
        self.playNum = 0
        self.length = len(self.wb.sheetnames)
        self.Activeplaylist = self.wb[self.wb.sheetnames[self.playNum]]

        pass

    def Next_Playlist(self):
        if self.playNum >= (self.length - 1):
            self.playNum = 0
            pass
        else:
            self.playNum += 1
            pass

        self.Activeplaylist = self.wb[self.wb.sheetnames[self.playNum]]
        pass

    def Prev_Playlist(self):
        if self.playNum <= 0:
            self.playNum = self.length - 1
            pass
        else:
            self.playNum -= 1
            pass

        self.Activeplaylist = self.wb[self.wb.sheetnames[self.playNum]]
        pass


class Track():
    def __init__(self):
        self.folderpath = os.path.dirname(os.path.abspath(__file__)) + "\Music\Tracks"
        self.ws = Playlist().Activeplaylist

        self.trackNum = 2

        while self.ws[]
        pass

    def Next_Track(self):
        self.trackNum += 1
        self.selectedCellAddr = 'B' + str(self.trackNum)

        if self.ws[self.selectedCellAddr].value == None:
            self.trackNum = 2
            pass
        else:
            pass

        #PLAYS MUSIC
        
        pass

    def Prev_Track(self):
        self.trackNum -= 1
        self.selectedCellAddr = 'B' + str(self.trackNum)

        if self.ws[self.selectedCellAddr].value == 'Track Title':
            self.trackNum = 2
            pass
        else:
            pass

        #PLAYS MUSIC
        
        pass

if __name__ == "__main__":
    root = tk.Tk()
    MainApplication(root).pack(side="top", fill="both", expand=False)
    root.mainloop()