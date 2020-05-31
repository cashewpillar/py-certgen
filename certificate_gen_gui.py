# Python stdlib imports
import time
import random
import tkinter as tk
from tkinter import ttk
from certificate_generator_functions import *

#TODO: organize code, separate windows into other files
win_home = window("Certificate Generator")
win_generating = window("Generating Certificates...")
win_generated = window("Certificates Generated!")
win_view = window("Generated Certificates")

# ============== Certificate Generator  ==============
win_home.minsize(width=560, height=250)

frm_home_title = tk.Frame(win_home)
frm_home_body = tk.Frame(win_home)
frm_home_nav = tk.Frame(win_home)
frm_home_title.pack(fill=tk.BOTH, expand=True)
frm_home_body.pack(fill=tk.BOTH, pady=25, expand=True)
frm_home_nav.pack(fill=tk.BOTH, pady=15, expand=True)

texts = ["Names file (.xlsx):", "Certificate file (.jpg/.png):", "Save generated certificates at:"]
commands = [openxl, openjpg, openfolder]
texts_btn = ["Open File", "Open File", "Select Folder"]
for i in range(3):
    label = tk.Label(frm_home_body, text=texts[i])
    label.grid(row=i, column=0, padx=5, pady=2.5, sticky="w")

    entry = tk.Entry(frm_home_body, width=50)
    entry.grid(row=i, column=1, padx=5, pady=2.5, sticky="ew")

    button = tk.Button(frm_home_body, width=10, text=texts_btn[i], command=commands[i])
    button.grid(row=i, column=2, padx=5, pady=2.5, sticky="w")

lbl_title = tk.Label(master=frm_home_title, text="Simple Certificate Generator", font=("Arial", 24))
btn_generate = tk.Button(master=frm_home_nav, width=27, height=2, text="Generate Certificates", command=generate)
btn_exit = tk.Button(master=frm_home_nav, width=27, height=2, text="Exit", command=exit)

lbl_title.pack(fill=tk.BOTH, side=tk.LEFT)
btn_generate.pack(side=tk.LEFT, padx=20, pady=5)
btn_exit.pack(side=tk.LEFT, padx=20, pady=5)

frm_home_body.rowconfigure([0, 1], weight=1, minsize=25)
frm_home_body.columnconfigure(0, weight=1, minsize=len(texts[1]))
frm_home_body.columnconfigure(1, weight=1, minsize=50)
frm_home_body.columnconfigure(2, weight=1, minsize=25)

#programming process helper
def close():
    for win in (win_generated, win_generating, win_home, win_view):
        win.destroy()
btn_close = tk.Button(frm_home_nav, width=5, height=2, text="X", command=close)
btn_close.pack(side=tk.LEFT, padx=20, pady=5)

# ============== Generating Certificates...  ==============
win_generating.minsize(width=360, height=150)
win_generating.maxsize(width=360, height=150)
pb = ttk.Progressbar(win_generating, length=320, mode="determinate", value=0)
pb.pack(pady=35)
lbl_progress = tk.Label(win_generating, text="0/N Certificates generated...")
lbl_progress.pack()
pb.start()
# for i in range(100):
#     time.sleep(1)
#     pb.step(1)

# ============== Certificates Generated!  ==============
win_generated.minsize(width=360, height=150)
win_generated.maxsize(width=360, height=150)

frm_success_title = tk.Frame(win_generated)
frm_success_nav = tk.Frame(win_generated)
frm_success_title.pack(padx=10, pady=20, fill=tk.BOTH)
frm_success_nav.pack(padx=20, pady=5, fill=tk.BOTH)

lbl_title = tk.Label(frm_success_title, text="Successfully generated __ certificates!", font=(16))
lbl_title.pack(padx=13, side=tk.LEFT)

texts = ['View Certificates', 'Generate Again', 'Exit']
commands = [view, generate_again, exit]
directions = ["w", "ew", "e"]
for i in range(3):
    button = tk.Button(frm_success_nav, width=len(max(texts)), text=texts[i], command=commands[i])
    button.grid(row=0, column=i, padx=2.5, pady=5, sticky=directions[i])

frm_success_nav.columnconfigure([0, 1, 2], weight=1, minsize=25)

# ============== View Certificates  ==============
win_view.geometry("1080x720")
import os, re, math
from PIL import Image
from PIL import ImageTk
from datetime import datetime

src_certificate = r'C:/python_files/certificate_generator'
certificates = [f for f in os.listdir(src_certificate) if re.match(r'.*.png', f)]

now = datetime.now()
dt_str = now.strftime("%Y%m%d%H%M%S")
thumbnail_folder = os.path.join(src_certificate, f"{dt_str}_certificate_thumnails")
os.mkdir(thumbnail_folder)

for i, certificate in enumerate(certificates):
    img_certificate = Image.open(os.path.join(src_certificate, certificate))
    img_certificate.thumbnail((240, 135))
    img_certificate.save(os.path.join(thumbnail_folder, f"thumb{i}.jpg"))

def onFrameConfigure(canvas):
    canvas.configure(scrollregion=canvas.bbox("all"))

canvas = tk.Canvas(win_view, borderwidth=1)
frm_thumbnails = tk.Frame(canvas)
vsb = tk.Scrollbar(win_view, command=canvas.yview)
canvas.configure(yscrollcommand=vsb.set)
vsb.pack(side=tk.RIGHT, fill=tk.Y)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
canvas.create_window((4,4), window=frm_thumbnails, anchor="nw")
frm_thumbnails.bind("<Configure>", lambda event, canvas=canvas: onFrameConfigure(canvas))

os.chdir(thumbnail_folder)
for i, certificate in enumerate(os.listdir(thumbnail_folder)):
    photo = ImageTk.PhotoImage(Image.open(certificate), master=win_view)
    btn = tk.Button(frm_thumbnails, width=240, height=135, image=photo)
    btn.image = photo
    btn.grid(row=math.floor(i/4), column=i%4, padx=10, pady=10)
