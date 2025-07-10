import fitz  # PyMuPDF

def list_images_with_size(pdf_path):
    doc = fitz.open(pdf_path)
    image_infos = []

    for page_number in range(len(doc)):
        page = doc[page_number]
        images = page.get_images(full=True)

        for img_index, img in enumerate(images):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            size_bytes = len(image_bytes)

            image_infos.append({
                "page": page_number + 1,
                "xref": xref,
                "format": image_ext,
                "size_bytes": size_bytes,
                "size_ko": size_bytes / 1024,
                "size_mo": size_bytes / (1024 * 1024)
            })

    # Affichage
    for i, info in enumerate(image_infos, 1):
        print(f"Image {i}:")
        print(f"  Page        : {info['page']}")
        print(f"  Format      : {info['format']}")
        print(f"  Taille      : {info['size_ko']:.2f} Ko ({info['size_mo']:.2f} Mo)")
        print(f"  Référence xref : {info['xref']}")
        print()

    print(f"Nombre total d'images : {len(image_infos)}")

# Exemple d'utilisation
list_images_with_size("C:\\Users\\dharlet\\Downloads\\Niji-MINISTERE JUSTICE-TRA LOT 2-memoire technique_sans_optim_avec_signet.pdf")
