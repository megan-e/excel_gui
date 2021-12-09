import tkinter as tk
from tkinter import *
from tkinter import filedialog


import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import pandas as pd
from matplotlib.ticker import ScalarFormatter
from setuptools import setup


#spot plating
#current limitations: workable with one sheet only
root = Tk()
root.geometry("400x400")

class drop_spot_plating():
    def drop_spot(self):
        df['raw_cells_1'] = df['No_of_cells.1'] / df['Amount Plated'] / df['Dilution']
        df['raw_cells_2'] = df['No_of_cells.2'] / df['Amount Plated'] / df['Dilution']
        df['raw_cells_3'] = df['No_of_cells.3'] / df['Amount Plated'] / df['Dilution']
        df['Average_CFU'] = (df.iloc[:, 10:13].sum(axis=1) / 3) * (df['Volume (mL)'] / df['Mass (g)'])
        df['Total_CFU'] = df['Average_CFU'].groupby(df['Type']).transform('mean')
        df['STDEV'] = df['Average_CFU'].groupby(df['Type']).transform('std')
        df['CV %'] = (df['STDEV'] / df['Total_CFU']) * 100
        new_df = df.iloc[::3, 0:17]
        global final_df
        final_df = new_df[['Condition', 'Total_CFU', 'STDEV', 'CV %']].copy()
        print(final_df)
        print(df)
        label['text'] = final_df

    def bar_graph(self):
        fig = plt.figure()
        x = final_df['Condition']
        y = final_df['Total_CFU']
        yerror = final_df['STDEV']
        ax = fig.add_subplot(111)
        plt.errorbar(x, y, yerr=yerror, fmt='o', color='black')
        plt.bar(x, y)
        plt.xlabel("Condition")
        plt.ylabel("Total Log of Number of Cells")
        ax.yaxis.set_major_formatter(mtick.FormatStrFormatter('%.2e'))
        label['text'] = plt.show()

    def analysis_spot_drop(self):
        global df
        global path
        path = filedialog.askopenfilename(initialdir="/", title="select a file")
        print(path)
        df = pd.read_excel(path)
        print(df)
        label['text'] = drop_spot_plating.drop_spot(self)

    def save_worksheet(self):
        writer = pd.ExcelWriter(path, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='RAW', index=False)
        final_df.to_excel(writer, sheet_name='SUMMARY', index=False)
        writer.save()
        label['text'] = "Saved!"
# ds = drop_spot_plating()
class spot_plating_time():
    def spot_time(self):
        df_spot['raw_cells_1'] = df_spot['No_of_cells.1'] / df_spot['Amount Plated'] / df_spot['Dilution']
        df_spot['raw_cells_2'] = df_spot['No_of_cells.2'] / df_spot['Amount Plated'] / df_spot['Dilution']
        df_spot['raw_cells_3'] = df_spot['No_of_cells.3'] / df_spot['Amount Plated'] / df_spot['Dilution']
        df_spot['Average_CFU'] = (df_spot.iloc[:, 11:14].sum(axis=1) / 3) * (df_spot['Volume (mL)'] / df_spot['Mass (g)'])
        df_spot['Total_CFU'] = df_spot['Average_CFU'].groupby(df_spot['Type']).transform('mean')
        df_spot['STDEV'] = df_spot['Average_CFU'].groupby(df_spot['Type']).transform('std')
        df_spot['CV %'] = (df_spot['STDEV'] / df_spot['Total_CFU']) * 100
        new_df = df_spot.iloc[::3, 0:18]
        global final_df_spot_time
        final_df_spot_time = new_df[['Condition', 'Time', 'Total_CFU', 'STDEV', 'CV %']].copy()
        print(final_df_spot_time)
        print(df_spot)
        label['text'] = final_df_spot_time

    def analysis_spot_time(self):
        global df_spot
        global path
        path = filedialog.askopenfilename(initialdir="/", title="select a file")
        print(path)
        df_spot = pd.read_excel(path)
        print(df_spot)
        label['text'] = spot_plating_time.spot_time(self)

    def spot_time_graph(self):
        global final_df_spot_time
        if clicked_spot.get() == 2:
            fig = plt.figure()
            x = final_df_spot_time.iloc[0::2, 1]
            y = final_df_spot_time.iloc[0::2, 2]
            y1 = final_df_spot_time.iloc[1::2, 2]
            yerror = final_df_spot_time.iloc[0::2, 3]
            y1error = final_df_spot_time.iloc[1::2, 3]
            ##set minimum spec##

            ax = fig.add_subplot(111)
            ax.set_yscale('log')
            plt.errorbar(x, y, yerr=yerror, fmt='o', color='black')
            plt.errorbar(x, y1, yerr=y1error, fmt='o', color='black')
            plt.plot(x, y, marker='o', label=final_df_spot_time.iloc[0, 0])
            plt.plot(x, y1, marker='o', label=final_df_spot_time.iloc[1, 0])
            plt.xlabel("Time")
            plt.ylabel("Total Log of Number of Cells")
            ax.yaxis.set_major_formatter(mtick.FormatStrFormatter('%.2e'))
            ax.xaxis.set_major_formatter(ScalarFormatter())
            plt.gca().set_xlim(left=0)
            plt.gca().set_ylim(top=10.00e+08)
            plt.legend(loc='upper right')
            plt.show()
        if clicked_spot.get() == 3:
            fig = plt.figure()
            x = final_df_spot_time.iloc[0::3, 1]
            y = final_df_spot_time.iloc[0::3, 2]
            y1 = final_df_spot_time.iloc[1::3, 2]
            y2 = final_df_spot_time.iloc[2::3, 2]
            yerror = final_df_spot_time.iloc[0::3, 3]
            y1error = final_df_spot_time.iloc[1::3, 3]
            y2error = final_df_spot_time.iloc[2::3, 3]

            ax = fig.add_subplot(111)
            ax.set_yscale('log')
            plt.errorbar(x, y, yerr=yerror, fmt='o', color='black')
            plt.errorbar(x, y1, yerr=y1error, fmt='o', color='black')
            plt.errorbar(x, y2, yerr=y2error, fmt='o', color='black')

            plt.plot(x, y, marker='o', label=final_df_spot_time.iloc[0, 0])
            plt.plot(x, y1, marker='o', label=final_df_spot_time.iloc[1, 0])
            plt.plot(x, y2, marker='o', label=final_df_spot_time.iloc[2, 0])
            plt.xlabel("Time")
            plt.ylabel("Total Log of Number of Cells")
            ax.yaxis.set_major_formatter(mtick.FormatStrFormatter('%.2e'))
            ax.xaxis.set_major_formatter(ScalarFormatter())
            plt.gca().set_xlim(left=0)
            plt.gca().set_ylim(top=10.00e+08)
            plt.legend(loc='upper right')
            plt.show()

        if clicked_spot.get() == 4:
            fig = plt.figure()
            x = final_df_spot_time.iloc[0::4, 1]
            y = final_df_spot_time.iloc[0::4, 2]
            y1 = final_df_spot_time.iloc[1::4, 2]
            y2 = final_df_spot_time.iloc[2::4, 2]
            y3 = final_df_spot_time.iloc[3::4, 2]

            yerror = final_df_spot_time.iloc[0::4, 3]
            y1error = final_df_spot_time.iloc[1::4, 3]
            y2error = final_df_spot_time.iloc[2::4, 3]
            y3error = final_df_spot_time.iloc[3::4, 3]

            ax = fig.add_subplot(111)
            ax.set_yscale('log')
            plt.errorbar(x, y, yerr=yerror, fmt='o', color='black')
            plt.errorbar(x, y1, yerr=y1error, fmt='o', color='black')
            plt.errorbar(x, y2, yerr=y2error, fmt='o', color='black')
            plt.errorbar(x, y3, yerr=y3error, fmt='o', color='black')

            plt.plot(x, y, marker='o', label=final_df_spot_time.iloc[0, 0])
            plt.plot(x, y1, marker='o', label=final_df_spot_time.iloc[1, 0])
            plt.plot(x, y2, marker='o', label=final_df_spot_time.iloc[2, 0])
            plt.plot(x, y3, marker='o', label=final_df_spot_time.iloc[3, 0])

            plt.xlabel("Time")
            plt.ylabel("Total Log of Number of Cells")
            ax.yaxis.set_major_formatter(mtick.FormatStrFormatter('%.2e'))
            ax.xaxis.set_major_formatter(ScalarFormatter())
            plt.gca().set_xlim(left=0)
            plt.gca().set_ylim(top=10.00e+08)
            plt.legend(loc='upper right')
            plt.show()

        elif clicked_spot.get() == 5:
            fig = plt.figure()
            x = final_df_spot_time.iloc[0::5, 1]
            y = final_df_spot_time.iloc[0::5, 2]
            y1 = final_df_spot_time.iloc[1::5, 2]
            y2 = final_df_spot_time.iloc[2::5, 2]
            y3 = final_df_spot_time.iloc[3::5, 2]
            y4 = final_df_spot_time.iloc[4::5, 2]
            yerror = final_df_spot_time.iloc[0::5, 3]
            y1error = final_df_spot_time.iloc[1::5, 3]
            y2error = final_df_spot_time.iloc[2::5, 3]
            y3error = final_df_spot_time.iloc[3::5, 3]
            y4error = final_df_spot_time.iloc[4::5, 3]


            ax = fig.add_subplot(111)
            ax.set_yscale('log')
            plt.errorbar(x, y, yerr=yerror, fmt='o', color='black')
            plt.errorbar(x, y1, yerr=y1error, fmt='o', color='black')
            plt.errorbar(x, y2, yerr=y2error, fmt='o', color='black')
            plt.errorbar(x, y3, yerr=y3error, fmt='o', color='black')
            plt.errorbar(x, y4, yerr=y4error, fmt='o', color='black')


            plt.plot(x, y, marker='o', label=final_df_spot_time.iloc[0, 0])
            plt.plot(x, y1, marker='o', label=final_df_spot_time.iloc[1, 0])
            plt.plot(x, y2, marker='o', label=final_df_spot_time.iloc[2, 0])
            plt.plot(x, y3, marker='o', label=final_df_spot_time.iloc[3, 0])
            plt.plot(x, y4, marker='o', label=final_df_spot_time.iloc[4, 0])

            plt.xlabel("Time")
            plt.ylabel("Total Log of Number of Cells")
            ax.yaxis.set_major_formatter(mtick.FormatStrFormatter('%.2e'))
            ax.xaxis.set_major_formatter(ScalarFormatter())
            plt.gca().set_xlim(left=0)
            plt.gca().set_ylim(top=10.00e+08)
            plt.legend(loc='upper right')
            plt.show()

    def save_worksheet(self):
        writer = pd.ExcelWriter(path, engine='xlsxwriter')
        df_spot.to_excel(writer, sheet_name='RAW', index=False)
        final_df_spot_time.to_excel(writer, sheet_name='SUMMARY', index=False)
        writer.save()
        label['text'] = "Saved!"

class spiral_plating():
    def spiral(self):
        dfs['Average_CFU'] = (dfs.iloc[:, 5:8].sum(axis=1) / 3) * (dfs['Volume (mL)'] / dfs['Mass (g)'])
        dfs['Total_CFU'] = dfs['Average_CFU'].groupby(dfs['Type']).transform('mean')
        dfs['STDEV'] = dfs['Average_CFU'].groupby(dfs['Type']).transform('std')
        dfs['CV %'] = (dfs['STDEV'] / dfs['Total_CFU']) * 100
        new_df = dfs.iloc[::3, 0:13]
        global df_spiral
        df_spiral = new_df[['Condition', 'Total_CFU', 'STDEV', 'CV %']].copy()
        print(df_spiral)
        label['text'] = df_spiral

    def analysis_spiral(self):
        global dfs
        global path
        path = filedialog.askopenfilename(initialdir="/", title="select a file")
        print(path)
        dfs = pd.read_excel(path)
        print(dfs)
        label['text'] = spiral_plating.spiral(self)

    def graph_spiral(self):
        fig = plt.figure()
        x = df_spiral['Condition']
        y = df_spiral['Total_CFU']
        yerror = df_spiral['STDEV']
        ax = fig.add_subplot(111)
        plt.errorbar(x, y, yerr=yerror, fmt='o', color='black')
        plt.bar(x, y)
        plt.xlabel("Condition")
        plt.ylabel("Total Log of Number of Cells")
        ax.yaxis.set_major_formatter(mtick.FormatStrFormatter('%.2e'))
        label['text'] = plt.show()

    def save_worksheet(self):
        writer = pd.ExcelWriter(path, engine='xlsxwriter')
        dfs.to_excel(writer, sheet_name='RAW', index=False)
        df_spiral.to_excel(writer, sheet_name='SUMMARY', index=False)
        writer.save()
        label['text'] = "Saved!"

class spiral_plating_time():
    def spiral(self):
        dff['Average_CFU'] = (dff.iloc[:, 6:9].sum(axis=1) / 3) * (dff['Volume (mL)'] / dff['Mass (g)'])
        dff['Total_CFU'] = dff['Average_CFU'].groupby(dff['Type']).transform('mean')
        dff['STDEV'] = dff['Average_CFU'].groupby(dff['Type']).transform('std')
        dff['CV %'] = (dff['STDEV'] / dff['Total_CFU']) * 100
        new_dff = dff.iloc[::3, 0:14]
        print(new_dff)
        global df_spiral_time
        df_spiral_time = new_dff[['Condition', 'Time', 'Total_CFU', 'STDEV', 'CV %']].copy()
        print(df_spiral_time)
        label['text'] = df_spiral_time

    def analysis_spiral(self):
        global dff
        global path
        path = filedialog.askopenfilename(initialdir="/", title="select a file")
        print(path)
        dff = pd.read_excel(path)
        print(dff)
        label['text'] = spiral_plating_time.spiral(self)

    def scatter_plot(self):
        global df_spiral_time
        if clicked.get() == 2:
            fig = plt.figure()
            x = df_spiral_time.iloc[0::2, 1]
            y = df_spiral_time.iloc[0::2, 2]
            y1 = df_spiral_time.iloc[1::2, 2]
            yerror = df_spiral_time.iloc[0::2, 3]
            y1error = df_spiral_time.iloc[1::2, 3]
            ##set minimum spec##


            ax = fig.add_subplot(111)
            ax.set_yscale('log')
            plt.errorbar(x, y, yerr=yerror, fmt='o', color='black')
            plt.errorbar(x, y1, yerr=y1error, fmt='o', color='black')
            plt.plot(x, y, marker='o', label=df_spiral_time.iloc[0, 0])
            plt.plot(x, y1, marker='o', label=df_spiral_time.iloc[1, 0])
            plt.xlabel("Time")
            plt.ylabel("Total Log of Number of Cells")
            ax.yaxis.set_major_formatter(mtick.FormatStrFormatter('%.2e'))
            ax.xaxis.set_major_formatter(ScalarFormatter())
            plt.gca().set_xlim(left=0)
            plt.gca().set_ylim(top=10.00e+08)
            plt.legend(loc='upper right')
            plt.show()
        if clicked.get() == 3:
            fig = plt.figure()
            x = df_spiral_time.iloc[0::3, 1]
            y = df_spiral_time.iloc[0::3, 2]
            y1 = df_spiral_time.iloc[1::3, 2]
            y2 = df_spiral_time.iloc[2::3, 2]
            yerror = df_spiral_time.iloc[0::3, 3]
            y1error = df_spiral_time.iloc[1::3, 3]
            y2error = df_spiral_time.iloc[2::3, 3]

            ax = fig.add_subplot(111)
            ax.set_yscale('log')
            plt.errorbar(x, y, yerr=yerror, fmt='o', color='black')
            plt.errorbar(x, y1, yerr=y1error, fmt='o', color='black')
            plt.errorbar(x, y2, yerr=y2error, fmt='o', color='black')

            plt.plot(x, y, marker='o', label=df_spiral_time.iloc[0, 0])
            plt.plot(x, y1, marker='o', label=df_spiral_time.iloc[1, 0])
            plt.plot(x, y2, marker='o', label=df_spiral_time.iloc[2, 0])
            plt.xlabel("Condition")
            plt.ylabel("Total Log of Number of Cells")
            ax.yaxis.set_major_formatter(mtick.FormatStrFormatter('%.2e'))
            ax.xaxis.set_major_formatter(ScalarFormatter())
            plt.gca().set_xlim(left=0)
            plt.gca().set_ylim(top=10.00e+08)
            plt.legend(loc='upper right')
            plt.show()

        if clicked.get() == 4:
            fig = plt.figure()
            x = df_spiral_time.iloc[0::4, 1]
            y = df_spiral_time.iloc[0::4, 2]
            y1 = df_spiral_time.iloc[1::4, 2]
            y2 = df_spiral_time.iloc[2::4, 2]
            y3 = df_spiral_time.iloc[3::4, 2]

            yerror = df_spiral_time.iloc[0::4, 3]
            y1error = df_spiral_time.iloc[1::4, 3]
            y2error = df_spiral_time.iloc[2::4, 3]
            y3error = df_spiral_time.iloc[3::4, 3]

            ax = fig.add_subplot(111)
            ax.set_yscale('log')
            plt.errorbar(x, y, yerr=yerror, fmt='o', color='black')
            plt.errorbar(x, y1, yerr=y1error, fmt='o', color='black')
            plt.errorbar(x, y2, yerr=y2error, fmt='o', color='black')
            plt.errorbar(x, y3, yerr=y3error, fmt='o', color='black')

            plt.plot(x, y, marker='o', label=df_spiral_time.iloc[0, 0])
            plt.plot(x, y1, marker='o', label=df_spiral_time.iloc[1, 0])
            plt.plot(x, y2, marker='o', label=df_spiral_time.iloc[2, 0])
            plt.plot(x, y3, marker='o', label=df_spiral_time.iloc[3, 0])

            plt.xlabel("Condition")
            plt.ylabel("Total Log of Number of Cells")
            ax.yaxis.set_major_formatter(mtick.FormatStrFormatter('%.2e'))
            ax.xaxis.set_major_formatter(ScalarFormatter())
            plt.gca().set_xlim(left=0)
            plt.gca().set_ylim(top=10.00e+08)
            plt.legend(loc='upper right')
            plt.show()

        elif clicked.get() == 5:
            fig = plt.figure()
            x = df_spiral_time.iloc[0::5, 1]
            y = df_spiral_time.iloc[0::5, 2]
            y1 = df_spiral_time.iloc[1::5, 2]
            y2 = df_spiral_time.iloc[2::5, 2]
            y3 = df_spiral_time.iloc[3::5, 2]
            y4 = df_spiral_time.iloc[4::5, 2]
            yerror = df_spiral_time.iloc[0::5, 3]
            y1error = df_spiral_time.iloc[1::5, 3]
            y2error = df_spiral_time.iloc[2::5, 3]
            y3error = df_spiral_time.iloc[3::5, 3]
            y4error = df_spiral_time.iloc[4::5, 3]


            ax = fig.add_subplot(111)
            ax.set_yscale('log')
            plt.errorbar(x, y, yerr=yerror, fmt='o', color='black')
            plt.errorbar(x, y1, yerr=y1error, fmt='o', color='black')
            plt.errorbar(x, y2, yerr=y2error, fmt='o', color='black')
            plt.errorbar(x, y3, yerr=y3error, fmt='o', color='black')
            plt.errorbar(x, y4, yerr=y4error, fmt='o', color='black')


            plt.plot(x, y, marker='o', label=df_spiral_time.iloc[0, 0])
            plt.plot(x, y1, marker='o', label=df_spiral_time.iloc[1, 0])
            plt.plot(x, y2, marker='o', label=df_spiral_time.iloc[2, 0])
            plt.plot(x, y3, marker='o', label=df_spiral_time.iloc[3, 0])
            plt.plot(x, y4, marker='o', label=df_spiral_time.iloc[4, 0])

            plt.xlabel("Condition")
            plt.ylabel("Total Log of Number of Cells")
            ax.yaxis.set_major_formatter(mtick.FormatStrFormatter('%.2e'))
            ax.xaxis.set_major_formatter(ScalarFormatter())
            plt.gca().set_xlim(left=0)
            plt.gca().set_ylim(top=10.00e+08)
            plt.legend(loc='upper right')
            plt.show()

        else:
            print("error")

        # fig = plt.figure()
        # x = final_df.iloc[::3, 1]
        # y = final_df.iloc[0::3, 2]
        # y1 = final_df.iloc[1::3, 2]
        # y2 = final_df.iloc[2::3, 2]
        #
        # yerror = df_spiral_time['STDEV']
        # ax = fig.add_subplot(111)
        # plt.errorbar(x, y, yerr=yerror, fmt='o', color='black')
        # plt.bar(x, y)
        # ax.yaxis.set_major_formatter(mtick.FormatStrFormatter('%.2e'))
        # label['text'] = plt.show()

    def save_worksheet(self):
        writer = pd.ExcelWriter(path, engine='xlsxwriter')
        dff.to_excel(writer, sheet_name='RAW', index=False)
        df_spiral_time.to_excel(writer, sheet_name='SUMMARY', index=False)
        writer.save()
        label['text'] = "Saved!"

button1 = tk.Button(root, text="spot/drop", font=40, command=lambda: drop_spot_plating().analysis_spot_drop())
button1.grid(row=0, column=0)

button2 = tk.Button(root, text="save file drop spot", font=40, command=lambda: drop_spot_plating().save_worksheet())
button2.grid(row=1, column=0)

button3 = tk.Button(root, text="bar graph spot/drop", font=40, command=lambda: drop_spot_plating().bar_graph())
button3.grid(row=2, column=0)

buttonx = tk.Button(root, text="spiral", font=40, command=lambda: spiral_plating().analysis_spiral())
buttonx.grid(row=0, column=5)

buttony = tk.Button(root, text="save file spiral", font=40, command=lambda: spiral_plating().save_worksheet())
buttony.grid(row=1, column=5)

button4 = tk.Button(root, text="bar graph spiral", font=40, command=lambda: spiral_plating().graph_spiral())
button4.grid(row=2, column=5)

button5 = tk.Button(root, text="spiral time", font=40, command=lambda: spiral_plating_time().analysis_spiral())
button5.grid(row=0, column=8)

button6 = tk.Button(root, text="save file spiral", font=40, command=lambda: spiral_plating_time().save_worksheet())
button6.grid(row=1, column=8)

button7 = tk.Button(root, text="spot time", font=40, command=lambda: spot_plating_time().analysis_spot_time())
button7.grid(row=0, column=9)

button8 = tk.Button(root, text="spot time save", font=40, command=lambda: spot_plating_time().save_worksheet())
button8.grid(row=1, column=9)



options = [
    "1",
    "2",
    "3",
    "4",
    "5",
]

clicked = IntVar()
clicked.set("Insert # of Samples to plot")
drop = OptionMenu(root, clicked, *options, command=lambda x: spiral_plating_time().scatter_plot())
drop.grid(row=2, column=8)

clicked_spot = IntVar()
clicked_spot.set("Insert # of Samples to plot")
drop_spot = OptionMenu(root, clicked_spot, *options, command=lambda x: spot_plating_time().spot_time_graph())
drop_spot.grid(row=2, column=9)

lower_frame = tk.Frame(root, bd=10)
lower_frame.place(relx=0.5, rely=.3, relwidth=.80, relheight=0.5, anchor='n')
label = tk.Label(lower_frame, font=("Helvetica", 20))
label.place(relwidth=1, relheight=1)

root.mainloop()

