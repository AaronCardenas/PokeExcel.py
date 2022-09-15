#import
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import requests
import io
pokeapi = "https://pokeapi.co/api/v2/pokemon/"


#creamos el exel
wb = Workbook()


#primera gen
Gen1 = wb.active
Gen1.title = "Gen 1"
Gen1.column_dimensions['A'].width=15
Gen1.column_dimensions['B'].width=15
Gen1.column_dimensions['C'].width=10
Gen1['A1'] = "Nombre"
Gen1['B1'] = "Imagen"
Gen1['C1'] = "Capturado"
Gen1['D1'] = "Pixel Art"
for i in range(10):
    call = pokeapi + str(i+1)
    res = requests.get(call).json()
    nombre = res['name'].capitalize()
    foto = res['sprites']['front_default']
    res_foto = requests.get(foto)
    bits_foto = io.BytesIO(res_foto.content)
    imagen= Image(bits_foto)
    Gen1.row_dimensions[i+2].height=70
    Gen1['A'+str(i+2)] = nombre
    Gen1.add_image(imagen,'B'+str(i+2))


Gen2 = wb.create_sheet("Gen 2")
Gen2.column_dimensions['A'].width=15
Gen2.column_dimensions['B'].width=15
Gen2['A1'] = "Nombre"
Gen2['B1'] = "Imagen"
Gen2['C1'] = "Capturado"
Gen2['D1'] = "Pixel Art"
for i in range(151,251):
    call = pokeapi + str(i+1)
    res = requests.get(call).json()
    nombre = res['name'].capitalize()
    foto = res['sprites']['front_default']
    res_foto = requests.get(foto)
    bits_foto = io.BytesIO(res_foto.content)
    imagen= Image(bits_foto)
    Gen2.row_dimensions[i-149].height=70
    Gen2['A'+str(i-149)] = nombre
    Gen2.add_image(imagen,'B'+str(i-149))
    

wb.save("pokedex.xlsx")

    

print("end")
