import PyPDF2
import io
from PIL import Image
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import os

def extract_images_from_pdf(pdf_path):
    """
    Extrait toutes les images d'un fichier PDF.
    
    Args:
        pdf_path (str): Chemin vers le fichier PDF
        
    Returns:
        list: Liste des images extraites sous forme de tuples (image_data, page_number)
    """
    images = []
    
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                
                if '/Resources' in page and '/XObject' in page['/Resources']:
                    x_objects = page['/Resources']['/XObject']
                    
                    for obj in x_objects:
                        if x_objects[obj]['/Subtype'] == '/Image':
                            try:
                                image_data = x_objects[obj].get_data()
                                images.append((image_data, page_num + 1))
                            except Exception as e:
                                print(f"Erreur lors de l'extraction de l'image: {e}")
                                
    except Exception as e:
        print(f"Erreur lors de la lecture du PDF: {e}")
        
    return images

def create_image_list_window(images):
    """
    Crée une fenêtre avec une liste des images extraites.
    
    Args:
        images (list): Liste des images extraites
    """
    root = tk.Tk()
    root.title("Images extraites du PDF")
    
    # Création du Treeview
    tree = ttk.Treeview(root, columns=("Page", "Taille"), show="headings")
    tree.heading("Page", text="Page")
    tree.heading("Taille", text="Taille")
    
    # Ajout des images à la liste
    for i, (image_data, page_num) in enumerate(images):
        try:
            image = Image.open(io.BytesIO(image_data))
            size = f"{image.size[0]}x{image.size[1]}"
            tree.insert("", "end", values=(f"Page {page_num}", size))
        except Exception as e:
            print(f"Erreur lors du traitement de l'image {i}: {e}")
    
    # Ajout d'une barre de défilement
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    
    # Placement des widgets
    tree.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    root.mainloop()

def main():
    """
    Fonction principale qui permet de sélectionner un fichier PDF et d'afficher ses images.
    """
    root = tk.Tk()
    root.withdraw()  # Cache la fenêtre principale
    
    pdf_path = filedialog.askopenfilename(
        title="Sélectionner un fichier PDF",
        filetypes=[("Fichiers PDF", "*.pdf")]
    )
    
    if pdf_path:
        images = extract_images_from_pdf(pdf_path)
        if images:
            create_image_list_window(images)
        else:
            print("Aucune image trouvée dans le PDF.")

if __name__ == "__main__":
    main() 