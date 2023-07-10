import bs4
import requests
from pathlib import Path
import aspose.words as aw
from pypdf import PdfReader 
from os import remove

def guardar_imagenes(imagenes):
    ruta = Path(Path.home(), 'Downloads')
    al_menos_una = False
    for i in imagenes:
        img = i['src']
        nombre = img.split('/')[-1]
        try:
            imagen = requests.get(img)
            f = open(Path(Path(ruta), nombre), 'wb')        
            f.write(imagen.content)
            f.close()
            al_menos_una = True
        except:
            pass

    return al_menos_una
        

def extraer_imagenes_web(pagina):
    resultado = requests.get(pagina)
    sopa = bs4.BeautifulSoup(resultado.text, 'html.parser')
    imagenes = sopa.find_all('img')
    print('Extrayendo imagenes...')
    al_menos_una = guardar_imagenes(imagenes)
    if al_menos_una:
        print('Imagenes guardadas en la carpeta de "Descargas"')
    else:
        print('Lo siento. No se pudo extraer ninguna imagen. Puede ser que la pagina no lo permita...')


def extrar_imagenes_pdf(ruta):
    #Eliminar comillas de la ruta
    if ruta.startswith('"') and ruta.endswith('"'):
        ruta = ruta[1:-1]
    
    # Verificar extension del archivo
    extensiones_permitidas = ['pdf', 'doc', 'docx']
    if ruta.split('.')[-1] not in extensiones_permitidas:
        print('El archivo debe ser un PDF o un documento WORD')
        return
    
    # Convertir archivo
    convertido = False
    if ruta.split('.')[-1] == 'pdf':
        archivo = PdfReader(ruta)
    else:
        doc = aw.Document(ruta)
        doc.save(ruta + '.pdf')
        archivo = PdfReader(ruta + '.pdf')
        convertido = True
    
    ruta_descargas = Path(Path.home(), 'Downloads')
    al_menos_una = False
    
    # Busquedade y descargas de imagenes 
    for page in range(len(archivo.pages)):
        page = archivo.pages[page]
        try:
            for i in page.images:
                with open(Path(ruta_descargas, i.name), 'wb') as f:
                    f.write(i.data)
                    al_menos_una = True
        except:
            pass
        
    # Comporbacion de una imagen descargada
    if al_menos_una:
        print('Imagenes guardadas en la carpeta de "Descargas"')
        print('Se recomienda cambiar los nombres a las imagenes extraidas.')
        input('Presione "Enter" para continuar el programa...')
    else:
        print('Lo siento. No se pudo extraer ninguna imagen')
    
    # Borrar pdf generado en caso necesario
    if convertido:
        remove(ruta + '.pdf')
        
        
def main():
    opcion = ''

    while opcion != '3':
        print('\n', '═' * 75, '\n', sep='')
        print('''
        Desea extraer imagenes de: 
        1) Pagina Web
        2) Archivo (PDF o WORD)
        3) Salir
        ''')
        opcion = input('\nOpcion: ')
        
        if opcion == '1':
            pagina = input('\nLink de la pagina web: ')
            extraer_imagenes_web(pagina)
        elif opcion == '2':
            archivo = input('\nRuta absoluta del archivo (click derecho sobre el archivo -> "Copiar como ruta de acceso"): ')
            extrar_imagenes_pdf(archivo)
        elif opcion == '3':
            print('\nFin del programa.')
        else:
            print('\nOpción incorrecta!\n')
            

if __name__ == '__main__':
    main()