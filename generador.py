import ezsheets
import re
FACTURAS = '1SrbMMIg5mhrdQpqRmlQtN4SQ3OG6yYBQ196z6yGxhBc'
PLANTILLA = '1LgOAnwIsoldnI1IgmVxYUoHfm7357nWukQ_IZnxzfg4'
CARPETA = '1kZijA4Ooy8k4BExztdTkGjTj9PbzPa0d'




def copiarPlantilla(titulo):
    ss = ezsheets.Spreadsheet(PLANTILLA)
    ss2 = ezsheets.createSpreadsheet(titulo)
    
    

    ss[0].copyTo(ss2)
    ss2[0].delete()

    return  ss2

def rellenerPlantilla(hoja_nueva, datos):
    celdas_a_cambiar = {
  
        'NOMBRECLIENTE' : 'D5',
        'CIF' : 'D6',
        'DIRECCION' : 'D7',

        'CP' : 'D8',
        'CIUDAD' : 'D8',
        'FECHA' : 'D11',
        'N_FACTURA' : 'D12',
        
        'DESCRIPCION1' : 'D16',
        'DESCRIPCION2' : 'D17',
        'DESCRIPCION3' : 'D18',
        'DESCRIPCION4' : 'D19',
        'DESCRIPCION5' : 'D20',
        
        'CANTIDAD1' : 'E16',
        'CANTIDAD2' : 'E17',
        'CANTIDAD3' : 'E18',
        'CANTIDAD4' : 'E19',
        'CANTIDAD5' : 'E20',
        
        'PRECIO1' : 'F16',
        'PRECIO2' : 'F17',
        'PRECIO3' : 'F18',
        'PRECIO4': 'F19',
        'PRECIO5' : 'F20',
        
        
        
        'SUBTOTAL1' : 'G16',
        'SUBTOTAL2' : 'G17',
        'SUBTOTAL3' : 'G18',
        'SUBTOTAL4' : 'G19',
        'SUBTOTAL5' : 'G20',
        'SUBTOTAL' : 'G21',
        
        'IVA' : 'G22',
        'TOTAL' : 'G23'
    }


    for cabeza, celda in celdas_a_cambiar.items():
        hoja_nueva[celda] = re.sub(r'%\w+', datos[cabeza], hoja_nueva[celda])



def leerFacturas():
    
    ss = ezsheets.Spreadsheet(FACTURAS)
    return ss.sheets[0]

def leerTitulo(factura):
    return str(factura[0]) + '-' + str(factura[1])

def generarPDF(spread):
    spread.downloadAsPDF()

def borrarCopia(spread):
    spread.delete()

def main():
    facturas = leerFacturas()
    cabecera = [factura.upper() for factura in facturas.getRows(stopRow=2)[0]]
    for factura in facturas.getRows(startRow=2):
        titulo = leerTitulo(factura)
        spread_nueva = copiarPlantilla(titulo)
        rellenerPlantilla(spread_nueva[0], dict(zip(cabecera,factura)))
        generarPDF(spread_nueva)
        borrarCopia(spread_nueva)

if __name__ == '__main__':
    main()