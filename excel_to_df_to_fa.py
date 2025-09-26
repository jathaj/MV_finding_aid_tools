import pandas as pd
import tkinter as tk
from tkinter import filedialog
import time


def open_file():
     filepath = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))
     if filepath:
          current_file.set(filepath)
          display_file_name = filepath.split("/")[-1]
          selected_file_label.config(text = f"Selected file:\n{display_file_name}")
          progress_msg.config(text="")

def process_file():
    progress_msg.config(text="Now processing...Please wait!")
    time.sleep(2)

    try:
        data = pd.ExcelFile(current_file.get())

        sheet_names = data.sheet_names

        collection_data = pd.read_excel(data, sheet_name=sheet_names[2])
        box_data = pd.read_excel(data, sheet_name=sheet_names[1])

        collection_data = collection_data.fillna('')
        box_data = box_data.fillna('')

        '''
        The following block iterates through the collection metadata
        and stores each item in a variable. These variables are then used
        to populate the HTML template below with information from the
        input spreadsheet.
        '''

        for index, row in collection_data.iterrows():
                recordid = row['<recordid>']
                repository = row['<repository>']
                unititle = row["<unittitle>"]
                origination = row["<origination>"]
                unitdate = row["<unitdate>"]
                geogname = row["<geogname>"]
                abstract = row["<abstract>"]
                physdesc = row["<physdesc>"]
                scopecontent = row["<scopecontent>"]
                index_terms = row["<index>"]
                custodhist = row["<custodhist>"]
                acqinfo = row["<acqinfo>"]
                unitid = row["<unitid>"]
                originalsloc = row["<originalsloc>"]
                accessrestrict = row["<accessrestrict>"]
                language = row["<languageset>"]
                author = row["<author>"]
                eventdatetime = row["<eventdatetime>"]

        '''
        The following code iterates through the box content worksheet in the input spreadsheet
        and builds up an HTML string that is then expanded in the larger HTML
        template below so that the final HTML document contains an ordered listing of each box
        with its contents.
        '''       

        contents_dict = {}

        for index, row in box_data.iterrows():
            if row["Box"] not in contents_dict.keys():
                contents_dict[row["Box"]] = [[row["Description"], row["Dates"], row["Container"], row["Condition"]]]
            else:
                contents_dict[row["Box"]].append([row["Description"], row["Dates"], row["Container"], row["Condition"]])


        content_str = "<h3>(Item Description | Dates | Container | Condition)</h3>\n"

        for key in contents_dict.keys():
            content_str += f"<h3>{key}</h3>\n \
                <ul>\n"
            for val in contents_dict[key]:
                item_str = ""
                for item in val:
                    if item != val[-1]:
                        item_str += f"{item} | "
                    else:
                        item_str += f"{item}"
                content_str += f"<li>{item_str}</li>\n"
            content_str += "</ul>\n"
                        
        '''
        The following code constructs the HTML document from the spreadsheet
        data. It then writes the HTML to a file.
        '''

        mvpl_img_url = "https://library.mountainview.gov/Project/Contents/Library/_gfx/cmn/mobile/mobile-logo.png"

        mvha_img_url = "https://www.mountainviewhistorical.org/wp-content/uploads/2022/09/mvha-logo-2023-header-1x-1024x164.png"

        institution = "mvpl" if custodhist == "MVPL" else "mvha"

        html_output = f"<!DOCTYPE html>\n \
                                <html lang=\"en\">\n \
                                    <head>\n \
                                        <meta charset=\"UTF-8\">\n \
                                        <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n \
                                            <meta http-equiv=\"X-UA-Compatible\" content=\"ie=edge\">\n \
                                        <title>{unititle}</title>\n \
                                        <link rel=\"stylesheet\" href=\"style.css\">\n \
                                    </head>\n \
                                    <body>\n \
                                        <img class=\"{institution}\" src=\"{mvpl_img_url if custodhist == "MVPL" else mvha_img_url}\">\n \
                                        <h1>Guide to the {unititle}</h1>\n \
                                            <p>Mountain View Public Library</p>\n  \
                                            <p>585 Franklin Street</p>\n \
                                            <p>Mountain View, CA 94041</p>\n \
                                            <p>Collection Number: {recordid}</p>\n \
                                        <hr>\n \
                                        <h2>Descriptive Summary</h2>\n \
                                        <h3>Title: <span class='interior'>{unititle}</span></h3>\n \
                                        <h3>Creator: <span class='interior'>{origination}</span></h3>\n \
                                        <h3>Dates: <span class='interior'>{unitdate}</span></h3>\n \
                                        <h3>Spatial Coverage: <span class='interior'>{geogname}</span></h3>\n \
                                        <h3>Extent: <span class='interior'>{physdesc}</span></h3>\n \
                                        <h3>Scope and Contents: <span class='interior'>{scopecontent}</span></h3>\n \
                                        <h3>Subjects and Indexing Terms: <span class='interior'>{index_terms}</span></h3>\n \
                                        <h3>Significance: <span class='interior'>{abstract}</span></h3>\n \
                                        <hr>\n \
                                        <h2>Administrative Details</h2>\n \
                                        <h3>Ownership: <span class='interior'>{custodhist}</span></h3>\n \
                                        <h3>Donor: <span class='interior'>{acqinfo}</span></h3>\n \
                                        <h3>Accession Number: <span class='interior'>{unitid}</span></h3>\n \
                                        <h3>Storage Location: <span class='interior'>{originalsloc}</span></h3>\n \
                                        <h3>Access Conditions: <span class='interior'>{accessrestrict}</span></h3>\n \
                                        <h3>Language: <span class='interior'>{language}</span></h3>\n \
                                        <hr>\n \
                                        <h2>Contents listing</h2>\n \
                                        {content_str} \
                                    <script src=\"index.js\"></script>\n \
                                    </body>\n \
                                </html>\n"

        with open(f"{unititle}.html", "w") as f:
            f.write(html_output)

        progress_msg.config(text=f"Finished processing!\n {unititle}.html is now available!")
    except:
        progress_msg.config(text=f"{current_file.get().split("/")[-1]} is not a valid file!\nTry a different file!")


#gui setup

app = tk.Tk()

app.title("MV Finding Aid Creation Tool")

current_file = tk.StringVar()

selected_file_label = tk.Label(app, text="No File Selected!")

load_file_btn = tk.Button(app, text="Load File", command=open_file)

process_file_btn = tk.Button(app, text="Process File", command=process_file)

progress_msg = tk.Label(app, text = "")


selected_file_label.grid(row=4, column=1, pady=10)

progress_msg.grid(row=6, column=1, pady=10)

load_file_btn.grid(row=3, column=1, pady=10)

process_file_btn.grid(row=5, column=1, pady=10)


app.mainloop()

'''
the follwing may be used in the future to support outputting 
EAD3 xml documents
'''
# xml = f"<?xml version="1.0" encoding="UTF-8"?> \
# <?xml-model href="schema/ead3.rng" type="application/xml" schematypens="http://relaxng.org/ns/structure/1.0"?> \
# <ead xmlns="http://ead3.archivists.org/schema/" audience="external"> <!-- EAD3 required element --> \
# <recordid>{}</recordid>


