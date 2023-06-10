import csv
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
import tkinter as tk
from tkinter import *
from tkinter import filedialog
import fitz
import customtkinter
from PIL import ImageTk, Image
import os
from typing import Union
from typing import Callable

WIDTH, HEIGHT = A4
MARGIN = 0
CELL_WIDTH = (WIDTH - MARGIN) / 2
CELL_HEIGHT = (HEIGHT - MARGIN) / 5
ROW_SPACING = 0 * cm

largeur = 1000
hauteur = 600

ma_variable = "oui"

class FloatSpinbox(customtkinter.CTkFrame):
    def __init__(self, *args,
                 width: int = 100,
                 height: int = 32,
                 step_size: Union[int, float] = 1,
                 command: Callable = None,
                 **kwargs):
        super().__init__(*args, width=width, height=height, **kwargs)

        self.step_size = step_size
        self.command = command

        self.configure(fg_color=("gray78", "gray28"))  # set frame color

        self.grid_columnconfigure((0, 2), weight=0)  # buttons don't expand
        self.grid_columnconfigure(1, weight=1)  # entry expands

        self.subtract_button = customtkinter.CTkButton(self, text="-", width=height-6, height=height-6,
                                                       command=self.subtract_button_callback)
        self.subtract_button.grid(row=0, column=0, padx=(3, 0), pady=3)

        self.entry = customtkinter.CTkEntry(self, width=width-(2*height), height=height-6, border_width=0)
        self.entry.grid(row=0, column=1, columnspan=1, padx=3, pady=3, sticky="ew")

        self.add_button = customtkinter.CTkButton(self, text="+", width=height-6, height=height-6,
                                                  command=self.add_button_callback)
        self.add_button.grid(row=0, column=2, padx=(0, 3), pady=3)

        # default value
        self.entry.insert(0, "0.0")

    def add_button_callback(self):
        if self.command is not None:
            self.command()
        try:
            value = float(self.entry.get()) + self.step_size
            self.entry.delete(0, "end")
            self.entry.insert(0, value)
        except ValueError:
            return

    def subtract_button_callback(self):
        if self.command is not None:
            self.command()
        try:
            value = float(self.entry.get()) - self.step_size
            self.entry.delete(0, "end")
            self.entry.insert(0, value)
        except ValueError:
            return

    def get(self) -> Union[float, None]:
        try:
            return float(self.entry.get())
        except ValueError:
            return None

    def set(self, value: float):
        self.entry.delete(0, "end")
        self.entry.insert(0, str(float(value)))
def clic_bouton():
    global csvfile
    page_w = float(spinbox_page_w.get()) * cm
    page_h = float(spinbox_page_h.get()) * cm
    box_line = float(spinbox_box_line.get())
    box_col = float(spinbox_box_column.get())
    marge = float(spinbox_marge.get())
    space = float(spinbox_spaceing_box.get())
    size_text = float(spinbox_size_text.get())
    left_space_text = float(spinbox_left_space_text.get())
    top_space_text = float(spinbox_top_space_text.get())
    nombres = champ_config.get().split(",")
    print(nombres)
    nombres_entiers = [int(nombre) for nombre in nombres]
    box_w = (page_w - (marge * 2) - space * (box_line-1)) / box_line
    box_h = (page_h - (marge * 2) - space * (box_col-1)) / box_col
    print(csvfile)
    with open(csvfile, newline='') as csvfiledata:
        reader = csv.reader(csvfiledata)
        c = canvas.Canvas(champ_texte.get(), pagesize=A4)
        c.setFont('Helvetica', size_text)
        lin_num = 0
        col_num = 0
        for row in reader:
            x = marge + (lin_num * box_w) + (lin_num * space)
            y = page_h - ((marge) + ((col_num -1) * box_h) + (col_num * space))
            c.rect(x, y - (box_h *2), box_w, box_h)
            for i in range(len(nombres_entiers)):
                c.drawString(x + left_space_text, y - box_h - ((i+1) * size_text) - top_space_text, row[nombres_entiers[i]])
            lin_num += 1
            if lin_num == box_line:
                col_num += 1
                lin_num = 0
            if col_num == box_col:
                c.showPage()
                col_num = 0
                lin_num = 0
        c.save()

def afficher_apercu_pdf():
    clic_bouton()
    chemin_pdf = champ_texte.get()
    doc = fitz.open(chemin_pdf)
    page = doc.load_page(0)
    pix = page.get_pixmap()
    image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    resized_image = image.resize((21*14, 30*14))
    photo = ImageTk.PhotoImage(resized_image)
    apercu_label.configure(image=photo)
    apercu_label.image = photo
    apercu_label.pack(side="right")
    os.remove(champ_texte.get())
    

def ouvrir_fichier():
    global csvfile
    csvfile = filedialog.askopenfilename()
    if csvfile:
        bouton_file.configure(text=os.path.basename(csvfile))
        with open(csvfile, 'r') as csvfiledata:
            reader = csv.reader(csvfiledata)
            en_tete = next(reader)
            colonnes_numerotees = [(i+1, colonne) for i, colonne in enumerate(en_tete)]
            colonnes_str = [f"{numero-1}: {colonne}" for numero, colonne in colonnes_numerotees]
            resultat = '\n'.join(colonnes_str)
        print(resultat)
        apercu_csv.configure(text=resultat)
        

def load():
    global csvfile
    print("load")
    save = filedialog.askopenfilename(filetypes=[("Fichiers texte", "*.AOF")])
    if save:
        with open(save, "r") as f:
            lignes = f.readlines()
            for ligne in lignes:
                nom, valeur = ligne.strip().split(": ")
                if nom == "Width Page":
                    spinbox_page_w.set(valeur)
                elif nom == "Height Page":
                    spinbox_page_h.set(valeur)
                elif nom == "Box per line":
                    spinbox_box_line.set(valeur)
                elif nom == "Box per column":
                    spinbox_box_column.set(valeur)
                elif nom == "Marge":
                    spinbox_marge.set(valeur)
                elif nom == "Box Spacing":
                    spinbox_spaceing_box.set(valeur)
                elif nom == "Size text":
                    spinbox_size_text.set(valeur)
                elif nom == "Text left space":
                    spinbox_left_space_text.set(valeur)
                elif nom == "Text top space":
                    spinbox_top_space_text.set(valeur)
                elif nom == "Information placement":
                    champ_config.delete(0, 999)
                    champ_config.insert(0, valeur)
                elif nom == "Outfile":
                    champ_texte.delete(0, 999)
                    champ_texte.insert(0, valeur)
                elif nom == "Csv file":
                    csvfile = os.path.basename(valeur)
        if csvfile:
            bouton_file.configure(text=os.path.basename(csvfile))
            with open(csvfile, 'r') as csvfiledata:
                reader = csv.reader(csvfiledata)
                en_tete = next(reader)
                colonnes_numerotees = [(i+1, colonne) for i, colonne in enumerate(en_tete)]
                colonnes_str = [f"{numero-1}: {colonne}" for numero, colonne in colonnes_numerotees]
                resultat = '\n'.join(colonnes_str)
            print(resultat)
            apercu_csv.configure(text=resultat)
        afficher_apercu_pdf()





def save():
    global csvfile
    fichier = filedialog.asksaveasfilename(defaultextension=".AOF", filetypes=[("SaveFile", "*.AOF")])
    if fichier:
        with open(fichier, "w") as f:
            filename = os.path.basename(csvfile)
            f.write("Csv file: {}\n".format(filename))
            f.write("Width Page: {}\n".format(spinbox_page_w.get()))
            f.write("Height Page: {}\n".format(spinbox_page_h.get()))
            f.write("Box per line: {}\n".format(spinbox_box_line.get()))
            f.write("Box per column: {}\n".format(spinbox_box_column.get()))
            f.write("Marge: {}\n".format(spinbox_marge.get()))
            f.write("Box Spacing: {}\n".format(spinbox_spaceing_box.get()))
            f.write("Size text: {}\n".format(spinbox_size_text.get()))
            f.write("Text left space: {}\n".format(spinbox_left_space_text.get()))
            f.write("Text top space: {}\n".format(spinbox_top_space_text.get()))
            f.write("Information placement: {}\n".format(champ_config.get()))
            f.write("Outfile: {}\n".format(champ_texte.get()))



global fenetre

color_theme = "green"
customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme(color_theme)

fenetre = customtkinter.CTk()
fenetre.title("AutoOrderForm")
fenetre.geometry(f"{largeur}x{hauteur}")


second_color = "transparent"

apercu_csv = customtkinter.CTkLabel(fenetre, text="\0", fg_color="transparent")
apercu_csv.pack(side="left")

apercu_label = customtkinter.CTkLabel(fenetre, text="\0", fg_color="transparent")
apercu_label.pack(side="right")

main_cadre = customtkinter.CTkFrame(fenetre)
# main_cadre.configure(fg_color="#00ff00")
main_cadre.pack()


cadre_bouton_file_csv = customtkinter.CTkFrame(main_cadre)
cadre_bouton_file_csv.pack(fill= X, padx=10,pady=10)
label_bouton_file_csv = customtkinter.CTkLabel(cadre_bouton_file_csv, text="Input File", fg_color=second_color)
label_bouton_file_csv.pack(side="left")
bouton_file = customtkinter.CTkButton(cadre_bouton_file_csv, text="CSV FILE", command=ouvrir_fichier)
bouton_file.pack(side="right")



cadre_page_w = customtkinter.CTkFrame(main_cadre)
cadre_page_w.pack(fill= X, padx=4,pady=2)
label_page_w = customtkinter.CTkLabel(cadre_page_w, text="Width Page:", fg_color=second_color)
label_page_w.pack(side="left")
spinbox_page_w = FloatSpinbox(cadre_page_w, width=150, step_size=1)
spinbox_page_w.set(21)
spinbox_page_w.pack(side="right")


cadre_page_h = customtkinter.CTkFrame(main_cadre)
cadre_page_h.pack(fill= X, padx=4,pady=2)
label_page_h = customtkinter.CTkLabel(cadre_page_h, text="Height Page:", fg_color=second_color)
label_page_h.pack(side="left")
spinbox_page_h = FloatSpinbox(cadre_page_h, width=150, step_size=1)
spinbox_page_h.set(29.7)
spinbox_page_h.pack(side="right")

cadre_box_line = customtkinter.CTkFrame(main_cadre)
cadre_box_line.pack(fill= X, padx=4,pady=2)
label_box_line = customtkinter.CTkLabel(cadre_box_line, text="Box per line:", fg_color=second_color)
label_box_line.pack(side="left")
spinbox_box_line = FloatSpinbox(cadre_box_line, width=150, step_size=1)
spinbox_box_line.set(1)
spinbox_box_line.pack(side="right")

cadre_box_column = customtkinter.CTkFrame(main_cadre)
cadre_box_column.pack(fill= X, padx=4,pady=2)
label_box_column = customtkinter.CTkLabel(cadre_box_column, text="Box per column:", fg_color=second_color)
label_box_column.pack(side="left")
spinbox_box_column = FloatSpinbox(cadre_box_column, width=150, step_size=1)
spinbox_box_column.set(1)
spinbox_box_column.pack(side="right")

cadre_marge = customtkinter.CTkFrame(main_cadre)
cadre_marge.pack(fill= X, padx=4,pady=2)
label_marge = customtkinter.CTkLabel(cadre_marge, text="Marge:", fg_color=second_color)
label_marge.pack(side="left")
spinbox_marge = FloatSpinbox(cadre_marge, width=150, step_size=1)
spinbox_marge.set(0)
spinbox_marge.pack(side="right")

cadre_spaceing_box = customtkinter.CTkFrame(main_cadre)
cadre_spaceing_box.pack(fill= X, padx=4,pady=2)
label_spaceing_box = customtkinter.CTkLabel(cadre_spaceing_box, text="Box Spacing:", fg_color=second_color)
label_spaceing_box.pack(side="left")
spinbox_spaceing_box = FloatSpinbox(cadre_spaceing_box, width=150, step_size=1)
spinbox_spaceing_box.set(0)
spinbox_spaceing_box.pack(side="right")

cadre_size_text = customtkinter.CTkFrame(main_cadre)
cadre_size_text.pack(fill= X, padx=4,pady=2)
label_size_text = customtkinter.CTkLabel(cadre_size_text, text="Size text:", fg_color=second_color)
label_size_text.pack(side="left")
spinbox_size_text = FloatSpinbox(cadre_size_text, width=150, step_size=1)
spinbox_size_text.set(15)
spinbox_size_text.pack(side="right")


cadre_left_space_text = customtkinter.CTkFrame(main_cadre)
cadre_left_space_text.pack(fill= X, padx=4,pady=2)
label_left_space_text = customtkinter.CTkLabel(cadre_left_space_text, text="Text left space:", fg_color=second_color)
label_left_space_text.pack(side="left")
spinbox_left_space_text = FloatSpinbox(cadre_left_space_text, width=150, step_size=1)
spinbox_left_space_text.set(35)
spinbox_left_space_text.set(0)
spinbox_left_space_text.pack(side="right")

cadre_top_space_text = customtkinter.CTkFrame(main_cadre)
cadre_top_space_text.pack(fill= X, padx=4,pady=2)
label_top_space_text = customtkinter.CTkLabel(cadre_top_space_text, text="Text top space:", fg_color=second_color)
label_top_space_text.pack(side="left")
spinbox_top_space_text = FloatSpinbox(cadre_top_space_text, width=150, step_size=1)
spinbox_top_space_text.set(0)
spinbox_top_space_text.pack(side="right")


cadre_config = customtkinter.CTkFrame(main_cadre)
cadre_config.pack(fill= X, padx=4,pady=2)
label_config = customtkinter.CTkLabel(cadre_config, text="(<-)Information placement (2,1,6,5...):", fg_color=second_color)
label_config.pack(side="left")
config_saisi = tk.StringVar()
champ_config = customtkinter.CTkEntry(cadre_config)
champ_config.insert(0, "1,2,12,13,14,15,16,4,17")
champ_config.pack(side="right")

cadre_texte = customtkinter.CTkFrame(main_cadre)
cadre_texte.pack(fill= X, padx=4,pady=2)
label_texte = customtkinter.CTkLabel(cadre_texte, text="Outfile:", fg_color=second_color)
label_texte.pack(side="left")
texte_saisi = tk.StringVar()
champ_texte = customtkinter.CTkEntry(cadre_texte)
champ_texte.insert(0, "order.pdf")
champ_texte.pack(side="right")





bouton = customtkinter.CTkButton(main_cadre, text="Preview", command=afficher_apercu_pdf, )
bouton.pack(fill= X, padx=4,pady=2)

bouton = customtkinter.CTkButton(main_cadre, text="Auto_Order_Form", command=clic_bouton, )
bouton.pack(fill= X, padx=4,pady=3)



cadre_backup = customtkinter.CTkFrame(fenetre)
cadre_backup.pack(side="bottom",)
bouton_load = customtkinter.CTkButton(cadre_backup, text="Load setting", command=load, )
bouton_load.pack(side="left", padx = 1)  
bouton_save = customtkinter.CTkButton(cadre_backup, text="Save setting", command=save, )
bouton_save.pack(side="right", padx = 1)

fenetre.mainloop()
