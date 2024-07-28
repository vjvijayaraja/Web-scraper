from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from bs4 import BeautifulSoup
import requests
from PyPDF2 import PdfReader

file_path = input("What would you like to name the file?: ")
doc = Document("Links to scrape.docx")
rels = doc.part.rels
iteration_counter = 0

try:
    with open(file_path, "w", encoding="utf-8") as file:
        for rel_id, rel in enumerate(rels, 1):
            if rels[rel].reltype == RT.HYPERLINK:
                iteration_counter += 1
                
                link_target = rels[rel].target_ref
                if link_target.endswith(".pdf"):
                    pdf_file_path = input("What would you like to name this pdf you are trying to download from " + link_target + " : ")
                    try:
                        response = requests.get(link_target)
                        response.raise_for_status()
                        if response.status_code == 200:
                            with open(pdf_file_path, "wb") as pdf_file:
                                pdf_file.write(response.content)
                                print(pdf_file_path + " successfully saved")

                                pdf_text = ""
                                with open(pdf_file_path, "rb") as pdf_file:
                                    pdf_reader = PdfReader(pdf_file)
                                    for page_num in range(len(pdf_reader.pages)):
                                        page = pdf_reader.pages[page_num]
                                        pdf_text += page.extract_text()
                                
                                file.write("\n\n=== PDF Content from " + link_target + " ===\n\n")
                                file.write(pdf_text)
                                print("PDF content appended to the existing text file")
                        else:
                            print("Failed to download file:", link_target)
                    except requests.RequestException as e:
                        print("An error occurred while downloading the PDF file from: ", link_target)
                        print("Error message:", str(e))
                        file.write("An error occured while downloading PDF file from this link: "+link_target +" \n\n")
                        file.write("Error message: "+str(e))
                else:
                    file.write("\nWebsite Link: " + link_target + "\n")
                    try:
                        link_body = requests.get(link_target, verify=False)
                        link_body.raise_for_status()
                        soup = BeautifulSoup(link_body.text, 'html.parser')
                        text_content = soup.get_text()
                        lines = text_content.splitlines()
                        for line in lines:
                            if line.strip():
                                file.write(line + "\n")
                            else:
                                file.write("\n")
                    except requests.RequestException as e:
                        print("An error occurred while retrieving content from link:", link_target)
                        print("Error message:", str(e))
                        file.write("Failed to retrieve content from this link.\n\n")
                        file.write("Error message: "+str(e))
    
    print("File", file_path, "has been saved")
except Exception as e:
    print("An error occurred while writing the file:", str(e))
finally:
    file.close()