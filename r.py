import os
import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as tb
from reportlab.lib.pagesizes import A4, A3, A5
from reportlab.pdfgen import canvas
import win32print
import win32api

class PrintApp:
    def __init__(self, root):
        self.root = root
        self.root.config(bg="black")
        self.root.wm_title("Application d'impression de documents")
        self.root.minsize(width=600, height=400)
        
        # Variables pour le chemin du dossier et la liste des documents
        self.folder_path = tk.StringVar()
        self.documents_listbox = tk.Listbox(root, selectmode=tk.SINGLE, width=80, height=20)
        self.documents_listbox.pack(pady=10)
        
        # Cadre principal pour organiser les widgets
        self.main_frame = tk.Frame(root, bg="black")
        self.main_frame.pack(expand=True, anchor='center')
        
        # Titre de l'application
        self.title = tk.Label(self.main_frame, text="Bienvenue à notre application d'impression", bg="black", fg="white", font=("Arial", 23))
        self.title.grid(column=0, row=0, columnspan=3, pady=10)
        
        # Étiquette et champ d'entrée pour le chemin du dossier
        self.lien_label = tk.Label(self.main_frame, fg="white", text="Entrer chemin dossier", bg="black", width=20, font=("Arial", 13))
        self.lien_label.grid(column=0, row=1, sticky="e", padx=10)
        
        self.lien_entry = tk.Entry(self.main_frame, bg="black", fg="white", width=28, font=("Arial", 13),
                                   selectbackground="black", selectforeground="white", insertbackground="white",
                                   textvariable=self.folder_path)
        self.lien_entry.grid(column=1, row=1, padx=(10, 2), pady=6, sticky="w")
        
        # Bouton pour chercher le dossier
        self.chercher_button = tb.Button(self.main_frame, text="Chercher", bootstyle="success-outline", width=15, command=self.load_folder)
        self.chercher_button.grid(column=2, row=1, padx=(2, 10), pady=10, sticky="w")
        
        # Options de format d'impression
        self.format_frame = tk.Frame(self.main_frame, bg="black")
        self.format_frame.grid(column=0, row=2, columnspan=3, pady=10)
        
        self.format_label = tk.Label(self.format_frame, text="Choisir le format d'impression:", bg="black", fg="white")
        self.format_label.pack(side=tk.LEFT, padx=5)
        
        self.format_options = ["A4", "A3", "A5"]
        self.format_var = tk.StringVar(value=self.format_options[0])
        self.format_menu = tk.OptionMenu(self.format_frame, self.format_var, *self.format_options)
        self.format_menu.pack(side=tk.LEFT)
        
        # Bouton d'impression
        self.print_button = tb.Button(self.main_frame, text="Imprimer", bootstyle="success-outline",  width=15, command=self.print_document)
        self.print_button.grid(column=0, row=3, columnspan=3, pady=10)
    
    def load_folder(self):
        # Ouvrir une boîte de dialogue pour sélectionner un dossier
        folder_selected = filedialog.askdirectory()
        if not os.path.isdir(folder_selected):
            # Afficher un avertissement si le dossier sélectionné n'est pas valide
            messagebox.showwarning("Avertissement", "Veuillez entrer un chemin de dossier valide.")
            return
        self.folder_path.set(folder_selected)
        self.update_document_list(folder_selected)
    
    def update_document_list(self, folder):
        # Supprimer tous les éléments de la Listbox
        self.documents_listbox.delete(0, tk.END)
        # Ajouter les fichiers PDF trouvés dans le dossier à la Listbox
        for filename in os.listdir(folder):
            if filename.endswith(".pdf"):  # Filtrer par type de fichier si nécessaire
                self.documents_listbox.insert(tk.END, filename)
    
    def print_document(self):
        folder_path = self.folder_path.get()
        if not os.path.isdir(folder_path):
            # Afficher un avertissement si le chemin du dossier n'est pas valide
            messagebox.showwarning("Avertissement", "Veuillez entrer un chemin de dossier valide.")
            return
        
        format_selected = self.format_var.get()
        page_size = self.get_page_size(format_selected)
        
        if not page_size:
            # Afficher une erreur si le format sélectionné n'est pas supporté
            messagebox.showerror("Erreur", f"Format d'impression {format_selected} non supporté.")
            return
        
        # Créer un fichier PDF récapitulatif
        recap_pdf_path = os.path.join(folder_path, "recap_impression.pdf")
        recap_canvas = canvas.Canvas(recap_pdf_path, pagesize=A4)
        recap_canvas.setFont("Helvetica", 12)
        recap_canvas.drawString(100, 800, "Récapitulatif des impressions")
        
        try:
            y_position = 750
            for filename in os.listdir(folder_path):
                if filename.endswith(".pdf"):
                    document_path = os.path.join(folder_path, filename)
                    
                    # Ajouter une entrée dans le fichier récapitulatif
                    recap_canvas.drawString(100, y_position, f"Document: {filename} - Format: {format_selected}")
                    y_position -= 20
                    
                    # Imprimer le fichier PDF en utilisant win32print
                    self.print_pdf(document_path)
                    
                    # Afficher un message d'information pour chaque document imprimé
                    self.show_info_message(document_path, format_selected)
                    
            recap_canvas.save()
            
        except Exception as e:
            # Afficher une erreur si quelque chose ne va pas lors de l'impression
            messagebox.showerror("Erreur", f"Une erreur s'est produite lors de l'impression : {str(e)}")
    
    def get_page_size(self, format):
        # Retourner la taille de page correspondant au format sélectionné
        if format == "A4":
            return A4
        elif format == "A3":
            return A3
        elif format == "A5":
            return A5
        else:
            return None
    
    def show_info_message(self, document_path, format):
        # Afficher un message d'information indiquant que le document a été imprimé
        messagebox.showinfo("Information", f"Document {os.path.basename(document_path)} imprimé au format {format}.")
    
    def print_pdf(self, pdf_path):
        # Obtenir le nom de l'imprimante par défaut
        printer_name = win32print.GetDefaultPrinter()
        
        # Envoyer le fichier PDF à l'imprimante
        win32api.ShellExecute(0, "print", pdf_path, None, ".", 0)

if __name__ == "__main__":
    root = tk.Tk()
    app = PrintApp(root)
    root.mainloop()








