import os
from tkinter import *
from tkinter import filedialog, messagebox
from gtts import gTTS
from playsound import playsound
from docx2pdf import convert
import aspose.words as aw
from PIL import Image
import time


def main():
    root = Tk()
    app = Toolkit(root)
    root.mainloop()


# ======================================================== TOOLKIT =====================================================
class Toolkit:

    def __init__(self, master):
        self.master = master
        self.master.title("Toolkit")
        self.master.config(bg="#041C32")
        self.master.geometry("700x450+0+0")
        self.master.resizable(False, False)

        self.title = Label(self.master, text="TOOLKIT", bg="#041C32", font=('arial', 10, 'bold'), fg="white")
        self.title.place(x=300, y=15)
        self.TextToSpeechConverter_button = Button(self.master, text="Text To Speech Converter",
                                                   command=self.TextToSpeechConverter_Window, bg='#ECDBBA',
                                                   fg='#191919', padx=25, pady=25)
        self.TextToSpeechConverter_button.place(x=80, y=100)

        self.DocumentConverter_button = Button(self.master, text="Document Converter",
                                               command=self.Document_Converter_Window, bg='#ECDBBA', fg='#191919',
                                               padx=25, pady=25)
        self.DocumentConverter_button.place(x=380, y=100)

        self.Compressor_button = Button(self.master, text="Compressor Converter", command=self.CompressorWindow,
                                        bg='#ECDBBA', fg='#191919', padx=25, pady=25)
        self.Compressor_button.place(x=80, y=300)

        self.MediaConverter_button = Button(self.master, text="Media Converter", command=self.Media_Converter_Window,
                                            bg='#ECDBBA', fg='#191919', padx=25, pady=25)
        self.MediaConverter_button.place(x=380, y=300)

        self.master.mainloop()

    def Document_Converter_Window(self):
        self.newWindow = Toplevel(self.master)
        self.app = DocumentConverter(self.newWindow)

    def TextToSpeechConverter_Window(self):
        self.newWindow = Toplevel(self.master)
        self.app = TextToSpeechConverter(self.newWindow)

    def CompressorWindow(self):
        self.newWindow = Toplevel(self.master)
        self.app = Compressor(self.newWindow)

    def Media_Converter_Window(self):
        self.newWindow = Toplevel(self.master)
        self.app = MediaConverter(self.newWindow)


# ==================================================== COMPRESSOR CONVERTER MAIN========================================

class Compressor:

    def __init__(self, master):
        self.master = master
        self.master.title("Compressor Converter")
        self.master.config(bg="#041C32")
        self.master.geometry("700x450+0+0")

        self.master.mainloop()


# ============================================== TXT TO SPEECH CONVERTER ===============================================
class TextToSpeechConverter:

    def __init__(self, master):
        self.master = master
        self.master.title("Document Converter")
        self.master.config(bg="#041C32")
        self.master.geometry("700x450+0+0")
        self.master.resizable(False, False)

        self.title = Label(self.master, text="TEXT TO SPEECH CONVERTER", bg="#041C32", font=('arial', 10, 'bold'),
                           fg="white")
        self.title.place(x=250, y=15)

        self.BrowseLabel = Label(self.master, text="Enter the text to be spoken", bg='#041C32', fg='white')
        self.BrowseLabel.place(x=80, y=100)

        self.BrowseLocationEntry = Entry(self.master, width=15, bd=4, font=10)
        self.BrowseLocationEntry.place(x=300, y=100)
        self.BrowseLocationEntry.insert(0, "")

        self.SubmitButton = Button(self.master, text="SUBMIT", width="15", command=self.play,
                                   bg='#ECDBBA', fg='#191919')
        self.SubmitButton.place(x=300, y=200)

        self.master.mainloop()

    def play(self):
        self.language = "en"

        self.myobj = gTTS(text=self.BrowseLocationEntry.get(), lang=self.language, slow=False)
        self.myobj.save("convert.wav")
        time.sleep(4)
        os.system("convert.wav")

    def Exit(self):
        self.master.destroy()

    def Reset(self):
        self.BrowseLocationEntry.set("")



class DocumentConverter:

    def __init__(self, master):
        self.master = master
        self.master.title("Document Converter")
        self.master.config(bg="#041C32")
        self.master.geometry("700x450+0+0")
        self.master.resizable(False, False)

        self.title = Label(self.master, text="DOCUMENT CONVERTER", bg="#041C32", font=('arial', 10, 'bold'), fg="white")
        self.title.place(x=250, y=15)

        self.convertDOCXtoPDF = Button(self.master, text="Convert DOCX to PDF", command=self.DOCXPDF_Window,
                                       bg='#ECDBBA',
                                       fg='#191919', padx=25, pady=25)
        self.convertDOCXtoPDF.place(x=80, y=100)

        self.convertPDFtoDOCX = Button(self.master, text="Convert PDF to DOCX", command=self.PDFDOCX_Window,
                                       bg='#ECDBBA',
                                       fg='#191919', padx=25, pady=25)
        self.convertPDFtoDOCX.place(x=380, y=100)

        self.convertDOCXtoTXT = Button(self.master, text="Convert DOCX to TXT", command=self.DOCXTXT_Window,
                                       bg='#ECDBBA',
                                       fg='#191919', padx=25, pady=25)
        self.convertDOCXtoTXT.place(x=80, y=200)

        self.convertTXTtoDOCX = Button(self.master, text="Convert TXT to DOCX", command=self.TXTDOCX_Window,
                                       bg='#ECDBBA',
                                       fg='#191919', padx=25, pady=25)
        self.convertTXTtoDOCX.place(x=380, y=200)

        self.convertTXTtoPDF = Button(self.master, text="Convert TXT to PDF", command=self.DOCXTXT_Window, bg='#ECDBBA',
                                      fg='#191919', padx=25, pady=25)
        self.convertTXTtoPDF.place(x=250, y=300)

        self.master.mainloop()

    def DOCXPDF_Window(self):
        self.newWindow = Toplevel(self.master)
        self.app = DOCXtoPDF_Converter(self.newWindow)

    def PDFDOCX_Window(self):
        self.newWindow = Toplevel(self.master)
        self.app = PDFtoDOCX_Converter(self.newWindow)

    def DOCXTXT_Window(self):
        self.newWindow = Toplevel(self.master)
        self.app = DOCXtoTXT_Converter(self.newWindow)

    def TXTPDF_Window(self):
        self.newWindow = Toplevel(self.master)
        self.app = TXTtoPDF_Converter(self.newWindow)

    def TXTDOCX_Window(self):
        self.newWindow = Toplevel(self.master)
        self.app = TXTtoDOCX_Converter(self.newWindow)


# ======================================================== DOCX TO PDF CONVERTER =======================================
class DOCXtoPDF_Converter:

    def __init__(self, master):
        self.master = master
        self.master.title("Document To PDF Converter")
        self.master.geometry("700x450+0+0")
        self.master.config(bg="#041C32")
        self.master.resizable(False, False)

        self.title = Label(self.master, text="DOCX TO PDF CONVERTER", bg="#041C32", font=('arial', 10, 'bold'), fg="white")
        self.title.place(x=250, y=15)

        self.BrowseFile = Label(self.master, text="Choose a File : ", bg="#041C32", fg="white")
        self.BrowseFile.place(x=100, y=200)

        self.ChooseFileButton = Button(self.master, text="Select", command=self.open_and_convert, bg='#ECDBBA',
                                       fg='#191919')
        self.ChooseFileButton.place(x=250, y=200)

        self.master.mainloop()

    def open_and_convert(self):
        self.file = filedialog.askopenfilename(filetypes=[("Word Files", "*docx")])
        convert(self.file, r'C:\Users\maarc\OneDrive\Desktop\Converted.pdf')
        messagebox.showinfo("Done", "File successfully converted")


# ======================================================== PDF TO DOCX CONVERTER =======================================
class PDFtoDOCX_Converter:

    def __init__(self, master):
        self.master = master
        self.master.title("PDF to DOCX Converter")
        self.master.geometry("700x450+0+0")
        self.master.config(bg="#041C32")
        self.master.resizable(False, False)

        self.title = Label(self.master, text="PDF TO DOCX CONVERTER", bg="#041C32", font=('arial', 10, 'bold'),
                           fg="white")
        self.title.place(x=250, y=15)

        self.BrowseFile = Label(self.master, text="Choose a File : ", bg="#041C32", fg="white")
        self.BrowseFile.place(x=100, y=200)

        self.ChooseFileButton = Button(self.master, text="Select", command=self.open_and_convert, bg='#ECDBBA',
                                       fg='#191919')
        self.ChooseFileButton.place(x=250, y=200)

        self.master.mainloop()

    def open_and_convert(self):
        self.file = filedialog.askopenfilename(filetypes=[("PDF Files", "*pdf")])
        self.doc = aw.Document(self.file)
        self.doc.save("pdf-to-doc-converted.docx")
        messagebox.showinfo("Done", "File successfully converted")


# ======================================================== DOC TO TXT CONVERTER ========================================
class DOCXtoTXT_Converter:

    def __init__(self, master):
        self.master = master
        self.master.title("DOCX To TXT Converter")
        self.master.geometry("700x450+0+0")
        self.master.config(bg="#041C32")
        self.master.resizable(False, False)

        self.title = Label(self.master, text="DOCX TO TXT CONVERTER", bg="#041C32", font=('arial', 10, 'bold'),
                           fg="white")
        self.title.place(x=250, y=15)

        self.BrowseFile = Label(self.master, text="Choose a File : ", bg="#041C32", fg="white")
        self.BrowseFile.place(x=100, y=200)

        self.ChooseFileButton = Button(self.master, text="Select", command=self.open_and_convert, bg='#ECDBBA',
                                       fg='#191919')
        self.ChooseFileButton.place(x=250, y=200)

        self.master.mainloop()

    def open_and_convert(self):
        self.file = filedialog.askopenfilename(filetypes=[("Word Files", "*docx")])
        self.document = aw.Document(self.file)
        self.document.save("result.txt")
        messagebox.showinfo("Done", "File converted successfully")


# ======================================================== TXT TO DOC CONVERTER ========================================
class TXTtoDOCX_Converter:

    def __init__(self, master):
        self.master = master
        self.master.title("TXT To DOCX Converter")
        self.master.geometry("700x450+0+0")
        self.master.config(bg="#041C32")
        self.master.resizable(False, False)

        self.title = Label(self.master, text="TXT TO DOCS CONVERTER", bg="#041C32", font=('arial', 10, 'bold'),
                           fg="white")
        self.title.place(x=250, y=15)

        self.BrowseFile = Label(self.master, text="Choose a File : ", bg="#041C32", fg="white")
        self.BrowseFile.place(x=100, y=200)

        self.ChooseFileButton = Button(self.master, text="Select", command=self.open_and_convert, bg='#ECDBBA',
                                       fg='#191919')
        self.ChooseFileButton.place(x=250, y=200)

        self.master.mainloop()

    def open_and_convert(self):
        self.file = filedialog.askopenfilename(filetypes=[("Text Files", "*txt")])
        self.document = aw.Document(self.file)
        self.document.save("result.docx")
        messagebox.showinfo("Done", "File converted successfully")


# =============================================== MEDIA CONVERTER MAIN =================================================
class MediaConverter:

    def __init__(self, master):
        self.master = master
        self.master.title("Media Converter")
        self.master.config(bg="#041C32")
        self.master.geometry("700x450+0+0")

        self.text = Label(self.master, text="MEDIA CONVERTER", bg="#041C32", fg="white")
        self.text.place(x=250, y=15)

        self.ImageConverter_button = Button(self.master, text="IMAGE CONVERTER", command=self.Image_Converter_Window,
                                            bg='#ECDBBA', fg='#191919', padx=25, pady=25)
        self.ImageConverter_button.place(x=80, y=100)

        self.master.mainloop()

    def Image_Converter_Window(self):
        self.newWindow = Toplevel(self.master)
        self.app = ImageConverter(self.newWindow)


# ======================================================== IMAGE CONVERTER MAIN ========================================
class ImageConverter:

    def __init__(self, master):
        self.master = master
        self.master.title("Image Converter")
        self.master.config(bg="#041C32")
        self.master.geometry("700x450+0+0")

        self.text = Label(self.master, text="IMAGE CONVERTER", bg="#041C32", fg="white")
        self.text.place(x=250, y=15)

        self.PNGtoJPG_button = Button(self.master, text="PNG to JPG Converter", command=self.PNGtoJPG_Window,
                                      bg='#ECDBBA', fg='#191919', padx=25, pady=25)
        self.PNGtoJPG_button.place(x=80, y=100)

        self.JPGtoPNG_button = Button(self.master, text="JPG to PNG Converter", command=self.JPGtoPNG_Window,
                                      bg='#ECDBBA', fg='#191919', padx=25, pady=25)
        self.JPGtoPNG_button.place(x=300, y=100)

        self.master.mainloop()

    def PNGtoJPG_Window(self):
        self.newWindow = Toplevel(self.master)
        self.app = PNGtoJPG_Converter(self.newWindow)

    def JPGtoPNG_Window(self):
        self.newWindow = Toplevel(self.master)
        self.app = JPGtoPNG_Converter(self.newWindow)


# ======================================================== PNG TO JPG CONVERTER ========================================
class PNGtoJPG_Converter:

    def __init__(self, master):
        self.master = master
        self.master.title("Image Converter")
        self.master.config(bg="#041C32")
        self.master.geometry("700x450+0+0")
        self.master.resizable(False, False)

        self.text = Label(self.master, text="PNG TO JPG CONVERTER", bg="#041C32", fg="white")
        self.text.place(x=250, y=15)

        self.convertToJPGButton = Button(self.master, text="Convert PNG to JPG", command=self.open_and_convert,
                                         bg='#ECDBBA', fg='#191919', padx=25, pady=25)
        self.convertToJPGButton.place(x=80, y=100)

        self.master.mainloop()

    def open_and_convert(self):
        self.import_file_path = filedialog.askopenfilename(filetypes=[("PNG Files", "*png")])
        self.imageJPG = Image.open(self.import_file_path).convert("RGB")
        self.export_file_path = filedialog.asksaveasfilename(defaultextension=".jpg")
        self.imageJPG.save(self.export_file_path)

        messagebox.showinfo("Done", "File converted successfully")


# ======================================================== JPG TO PNG CONVERTER ========================================
class JPGtoPNG_Converter:

    def __init__(self, master):
        self.master = master
        self.master.title("JPG to PNG Converter")
        self.master.config(bg="#041C32")
        self.master.geometry("700x450+0+0")
        self.master.resizable(False, False)

        self.text = Label(self.master, text="JPG TO PNG CONVERTER", bg="#041C32", fg="white")
        self.text.place(x=250, y=15)

        self.convertToJPGButton = Button(self.master, text="Convert JPG to PNG", command=self.open_and_convert,
                                         bg='#ECDBBA', fg='#191919', padx=25, pady=25)
        self.convertToJPGButton.place(x=80, y=100)

        self.master.mainloop()

    def open_and_convert(self):
        self.import_file_path = filedialog.askopenfilename(filetypes=[("JPG Files", "*jpg")])
        self.imageJPG = Image.open(self.import_file_path)
        self.export_file_path = filedialog.asksaveasfilename(defaultextension=".png")
        self.imageJPG.save(self.export_file_path)

        messagebox.showinfo("Done", "File converted successfully")


if __name__ == "__main__":
    main()
