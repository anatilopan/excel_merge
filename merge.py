import glob
import os
import sys
import time
from os.path import abspath, dirname

import click
import pandas as pd
import PySimpleGUI as sg


def verify_path(ctx, param, value):
    if value:
        if not os.path.exists(value):
            raise click.BadParameter("Path does not exist: {}".format(value))
    return value


def check_path(value):
    if value:
        if not os.path.exists(value):
            return False
    return True


DIRPATTERN = "bt"
FILEPATTERN = "*.xlsx"  # Adjust the pattern based on your file naming conventions
OUT_FILE = "concatenated_data.xlsx"
PRG_BAR_MAX = 100
NO_OF_DIRS = 0
NO_OF_FILES = 0


def concatenate_excel_files_with_header(
    parent_dir, directory_pattern, file_pattern, output_file, window=None
):
    # Get a list of all subdirectories matching the pattern
    subdirectorys = [
        directory
        for directory in os.listdir(parent_dir)
        if os.path.isdir(os.path.join(parent_dir, directory))
        and directory_pattern in directory
    ]
    if len(subdirectorys) == 0:
        subdirectorys.append(os.getcwd())
    if window:
        progress_bar_dir = window.window["-PROGDIR-"]
        progress_label_dir = window.window["-PROGDIRLABEL-"]
        progress_bar_files = window.window["-PROGFILES-"]
        progress_label_files = window.window["-PROGFILESLABEL-"]
    # Initialize DataFrames to store concatenated data and header data
    all_data = []
    total_progress_dir = len(subdirectorys)
    current_progress_dir = 0

    # Iterate through subdirectories
    for subdirectory in subdirectorys:
        if window:
            progress_label_dir.update(f"{current_progress_dir} / {total_progress_dir}")
            window.refresh()

        print(f"should be here... {current_progress_dir}/{total_progress_dir}")

        directory_path = os.path.join(parent_dir, subdirectory)

        # Find all Excel files matching the pattern in the subdirectory
        excel_files = glob.glob(os.path.join(directory_path, file_pattern))

        total_progress_files = len(excel_files)
        current_progress_files = 0
        # Iterate through Excel files
        for file in excel_files:
            if window:
                progress_label_files.update(
                    f"{current_progress_files} / {total_progress_files}"
                )
                window.refresh()
            # Read the entire Excel file
            df = pd.read_excel(file)

            # Separate the first 8 rows as header
            header = df.head(7)
            columns_for_header = {
                "HEADER DATA: Sales agreement ref/id": header.iloc[0, 1],
                "HEADER DATA: Customer account": header.iloc[1, 1],
                "HEADER DATA: Client name": header.iloc[2, 1],
                "HEADER DATA: Name of TP account manager": header.iloc[3, 1],
                "HEADER DATA: Name of client contact": header.iloc[4, 1],
                "HEADER DATA: Customer reference": header.iloc[1, 3],
                "HEADER DATA: Expiration date": header.iloc[2, 3],
                "HEADER DATA: Payment terms": header.iloc[3, 3],
                "HEADER DATA: Invoice period": header.iloc[4, 3],
                "HEADER DATA: Confirmation date": header.iloc[2, 5],
            }
            # Skip the first 8 rows and concatenate the rest into all_data
            data = df.iloc[6:]
            new_header = data.iloc[0].tolist()
            data = data[1:]
            data.columns = new_header
            for key, value in columns_for_header.items():
                data[key] = value

            for index, row in data.iterrows():
                pass
            all_data.append(data)

            time.sleep(2)
            current_progress_files += 1
            if window:
                progress_bar_files.update(
                    (current_progress_files / total_progress_files) * 100
                )
                progress_label_files.update(
                    f"{current_progress_files} / {total_progress_files}"
                )
                window.refresh()

        current_progress_dir += 1
        if window:
            progress_bar_dir.update((current_progress_dir / total_progress_dir) * 100)
            progress_label_dir.update(f"{current_progress_dir} / {total_progress_dir}")
            window.refresh()

    # Concatenate all subdirectory DataFrames into a single DataFrame
    if all_data:
        result = pd.concat(all_data, ignore_index=True)

        # Save the concatenated DataFrame to a new Excel file
        result.to_excel(output_file, index=False)
        print(f"Concatenated data saved to {output_file}")
    else:
        result = "No Excel files found matching the pattern."
    return result


# # concatenate_excel_files(parent_directory, DIRPATTERN, FILEPATTERN, output_excel)
# concatenate_excel_files_with_header(
#     parent_directory, DIRPATTERN, FILEPATTERN, output_excel
# )


# @click.command()
# @click.option(
#     "-p", "--file-pattern", help="UNIX style pattern to use while searching for files"
# )
# @click.option(
#     "-P",
#     "--dir-pattern",
#     help="UNIX style pattern to use while searching for directories",
# )
# @click.option("-O", "--output-dir", callback=verify_path, help="The output directory.")
# @click.option("-o", "--output-file", help="Output file name")
# @click.option("-c", "--cli-mode", help="no gui")
# @click.argument(
#     "path",
#     metavar="PATH",
#     callback=verify_path,
#     required=False,
# )
def create_win():
    layout = [
        [sg.Text("Selectează folderul principal te rog:")],
        [
            sg.Text("Parent directory:", (13, 1)),
            sg.In(key="-IN-PATH-", size=(60, 1), disabled=True, enable_events=True),
            sg.FolderBrowse(),
        ],
        [
            sg.Text("File pattern:", (13, 1)),
            sg.Input(FILEPATTERN, key="-FPATTERN-", size=(13, 1), disabled=True),
            sg.Push(),
            sg.Text("Directory pattern:"),
            sg.Input(DIRPATTERN, key="-DPATTERN-", size=(13, 1), disabled=True),
        ],
        [
            sg.Text("Output directory:", (13, 1)),
            sg.In(key="-OUT-PATH-", size=(60, 1), disabled=True),
            sg.FolderBrowse(target=(3, 1)),
        ],
        [
            sg.Text("Output file name:", (13, 1)),
            sg.In(OUT_FILE, key="-FILENAME-", size=(20, 1)),
            sg.HSep(),
            sg.Text("v. 1.0.0"),
            sg.HSep(),
        ],
        [
            sg.Push(),
            sg.Open("Start"),
            sg.Cancel("Cancel"),
        ],
    ]

    return sg.Window("Merge Billing Templates", layout)


class Window_Obj:
    def __init__(self, layout=None):
        self.layout = layout
        self.window = self.make_process_window()

    def make_process_window(self):
        if self.layout:
            layout = self.layout
        else:
            layout = [
                [sg.Push(), sg.Text("Progress meter"), sg.Push()],
                [
                    sg.Push(),
                    sg.Text("Folder(s):"),
                    sg.ProgressBar(100, orientation="h", size=(20, 2), key="-PROGDIR-"),
                    sg.Text("-/-", key="-PROGDIRLABEL-"),
                ],
                [
                    sg.Text("Files:"),
                    sg.ProgressBar(
                        100, orientation="h", size=(23, 2), key="-PROGFILES-"
                    ),
                    sg.Text("-/-", key="-PROGFILESLABEL-"),
                ],
            ]
        return sg.Window("Window Title", layout, finalize=True)

    def refresh(self):
        self.window.refresh()

    def read(self, **kwargs):
        events, values = self.window.read(**kwargs)
        return events, values


def main(**kwargs):
    if kwargs.get("path"):
        parent_directory = kwargs.get("path")
    else:
        parent_directory = os.getcwd()
        print(f"Runnin in (not set) {parent_directory}")
    if kwargs.get("cli-mode"):
        ## RUN WITHOUT GUI
        if kwargs.get("file-pattern"):
            print(kwargs.get("file-pattern"))

    sg.theme("SystemDefault")

    progress_win_active = False

    # layout_output = [
    #     [sg.Text("Doing stuff...")],
    #     [sg.Output(size=(88, 20), font="Courier 10")],
    #     [sg.Cancel()],
    # ]

    window = create_win()
    result = None
    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, "Exit") or event == "Cancel":
            break
        elif event == "-IN-PATH-":
            parent_directory = values["-IN-PATH-"]
            checkpath = check_path(values["-IN-PATH-"])
            if not checkpath:
                sg.popup(f"The path '{parent_directory}' does not exist!")
            else:
                window["-OUT-PATH-"].update(value=values["-IN-PATH-"])
        elif event == "Start" and not progress_win_active:
            parent_dir = values["-IN-PATH-"]
            output_dir = values["-OUT-PATH-"]
            dir_patt = values["-DPATTERN-"]
            file_patt = values["-FPATTERN-"]
            out_file_name = values["-FILENAME-"]

            if not parent_dir or not output_dir:
                sg.popup(
                    f"Please check that you selected a parrent directory and an output directory."
                )
            elif not dir_patt or not file_patt:
                sg.popup(
                    f"Please check that you entered a pattern for files and directories."
                )
            elif not out_file_name:
                sg.popup(f"Please check that you set an file name for the output.")
            else:
                out_file = os.path.join(output_dir, out_file_name).replace("/", "\\")
                window.Hide()
                progress_win = Window_Obj()
                ev, vals = progress_win.read(timeout=100)
                if ev == sg.WIN_CLOSED or ev == "Exit":
                    progress_win.Close()
                    progress_win_active = False
                    window.UnHide()
                result = concatenate_excel_files_with_header(
                    parent_dir,
                    dir_patt,
                    file_patt,
                    out_file,
                    progress_win,
                )
                if type(result) is str:
                    sg.popup(result)
                    progress_win.window.Close()
                    progress_win_active = False
                    window.UnHide()
                elif isinstance(result, pd.DataFrame):
                    progress_win.window.Close()
                    sg.popup("Done!", title="Completed!")
                    progress_win_active = False
                    window.UnHide()

                ev, vals = progress_win.read(timeout=100)
                if ev == sg.WIN_CLOSED or ev == "Exit":
                    progress_win.window.Close()
                    progress_win_active = False
                    window.UnHide()

    window.close()


if __name__ == "__main__":
    main()
