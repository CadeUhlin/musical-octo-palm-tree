import pandas as pd
from tkinter import Tk
import tkinter.messagebox
from tkinter.filedialog import askopenfilename

class REDBOX():
    def __init__(self) -> None:
        pass

    def main(self):
        Tk().withdraw()

        tkinter.messagebox.showinfo('WARNING', 'Please select the Prism file first')
        while True:
            prism = askopenfilename()
            if prism.find('prism') >= 0 or prism.find('Prism') >= 0 or prism.find('PRISM') >= 0:
                break

        tkinter.messagebox.showinfo('WARNING', 'Please select the RedBox file second')
        while True:
            Redbox = askopenfilename()
            if Redbox.find('redbox') >= 0 or Redbox.find('Redbox') >= 0 or Redbox.find('REDBOX') >= 0:
                break

        self.prism_df = pd.read_excel(prism)
        self.Redbox_df = pd.read_excel(Redbox)

        tkinter.messagebox.showinfo('WARNING', 'Please select specific user files. Hit cancel when done.')
        while True:
            spec_user = askopenfilename()
            if spec_user == '':
                break
            else:
                spec_user_df = pd.read_excel(spec_user)
                spec_numbers = spec_user_df['Phone number']
                p_numbers = self.prism_df['Called Digits']
                for number in spec_numbers:
                    if number in p_numbers:
                        index = p_numbers.index(number)
                        spec_user_df['Calling Label'][index] = spec_user_df['Agent Name'][0]

        date = []
        time = []
        for value in self.Redbox_df['Call Start Time']:
            space_index = str(value).find(' ')
            temp = str(value)[:space_index]
            date.append(temp)
            temp = str(value)[space_index+1:]
            time.append(temp)

        self.Redbox_df['Call Start Time'] = date
        self.Redbox_df.rename(columns={'Call Start Time':'Call Start Date'}, inplace=True)
        self.Redbox_df.insert(1, 'Call Start Time', time)

        combined_times = []
        for index in range(len(self.prism_df)):
            time = self.prism_df['Time'][index]
            try:
                time_hour = int(time[:2])
            except SyntaxError:
                time_hour = int(time[1:2])
            try:
                time_minute = int(time[3:5])
            except SyntaxError:
                time_minute = int(time[4:5])
            try:
                time_sec = int(time[-2:])
            except SyntaxError:
                time_sec = int(time[-1:])

            ring_time = self.prism_df['Ring Time'][index]
            try:
                ring_time_second = int(ring_time[-2:])
            except SyntaxError:
                ring_time_second = int(ring_time[:-1])

            sec = time_sec
            minute = time_minute
            hour = time_hour
            if time_sec + ring_time_second > 60:
                sec = (time_sec + ring_time_second)%60
                if minute + 1 >= 60:
                    minute = 0
                    hour = hour + 1
                else:
                    minute = minute + 1
            else:
                sec = time_sec + ring_time_second
            
            combined_time = None
            if hour < 10 and minute < 10 and sec <10:
                combined_time = f'0{hour}:0{minute}:0{sec}'
            elif hour and minute < 10:
                combined_time = f'0{hour}:0{minute}:{sec}'
            elif minute <10 and sec < 10:
                combined_time = f'{hour}:0{minute}:0{sec}'
            elif hour < 10 and sec < 10:
                combined_time = f'0{hour}:{minute}:0{sec}'
            elif hour < 10:
                combined_time = f'0{hour}:{minute}:{sec}'
            elif minute < 10:
                combined_time = f'{hour}:0{minute}:{sec}'
            elif sec < 10:
                combined_time = f'{hour}:{minute}:0{sec}'
            else:
                combined_time = f'{hour}:{minute}:{sec}'

            combined_times.append(combined_time)
        self.prism_df['Time'] = combined_times

        self.numbers = [None]*len(self.Redbox_df['Extension'])
        self.group = [None]*len(self.Redbox_df['Group'])
        gcf = self.GCF(len(self.Redbox_df))
        window = int(len(self.Redbox_df)/gcf)
        count = 1

        while count <= gcf:
            self.combine(window, count)
            count = count + 1

        self.Redbox_df['Extension'] = self.numbers
        self.Redbox_df['Group'] = self.group

        writer = pd.ExcelWriter('Redbox.xlsx')
        self.Redbox_df.to_excel(writer, sheet_name='01-10 June 2021',index=False)
        writer.save()

    def GCF(self, gcf):
        largest_divisor = 0
        for i in range(2, gcf):
            if gcf % i == 0 and gcf % i <= 100:
                largest_divisor = i 
        return largest_divisor

    def combine(self, window, count):
        collection = []
        call_start_time = self.Redbox_df.loc[:, 'Call Start Time']
        prism_time = self.prism_df.loc[:, 'Time']
        prism_digits = self.prism_df.loc[:, 'Calling Digits']
        prism_label= self.prism_df.loc[:, 'Called Label']

        for i in range((window * (count -1)), (window * count)):
            r_time = call_start_time[i]
            r_head = r_time[:6]
            try:
                r_sec = int(r_time[-2:])
            except SyntaxError:
                r_sec = int(r_time[-1:])
            except ValueError:
                r_sec = int(r_time[-1:])
            for j in range(len(prism_time)):
                p_time = prism_time[j]
                p_head = p_time[:6]
                try:
                    p_sec = int(p_time[-2:])
                except SyntaxError:
                    p_sec = int(p_time[-1:])
                except ValueError:
                    p_sec = int(p_time[-1:])
                if j not in collection and r_head == p_head and abs(int(r_sec - p_sec)) < 3:
                    self.numbers[i] = prism_digits[j]
                    self.group[i] = prism_label[j]
                    collection.append(j)
                    print('***')
                    break
            print('***', i)

if __name__ == '__main__':
    redbox = REDBOX()
    redbox.main()