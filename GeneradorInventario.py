from Utils import headers
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
while True:
    print("""
        ░█████╗░░█████╗░███╗░░░███╗███████╗██████╗░  ░██████╗░█████╗░████████╗
        ██╔══██╗██╔══██╗████╗░████║██╔════╝██╔══██╗  ██╔════╝██╔══██╗╚══██╔══╝
        ██║░░╚═╝██║░░██║██╔████╔██║█████╗░░██████╔╝  ╚█████╗░███████║░░░██║░░░
        ██║░░██╗██║░░██║██║╚██╔╝██║██╔══╝░░██╔══██╗  ░╚═══██╗██╔══██║░░░██║░░░
        ╚█████╔╝╚█████╔╝██║░╚═╝░██║███████╗██║░░██║  ██████╔╝██║░░██║░░░██║░░░
        ░╚════╝░░╚════╝░╚═╝░░░░░╚═╝╚══════╝╚═╝░░╚═╝  ╚═════╝░╚═╝░░╚═╝░░░╚═╝░░░
          """)
    print("            GENERADOR DE EXISTENCIAS ALMACENES LLENO SAT Y JDC           ")
    print("                              VERSIÓN 9.0                                ")
    print("                         ING. ERICK CASANOVA                             ")
    print("                           ÁREA DE SISTEMAS                              ")
    print("   |-------------------------------------------------------------------|")
    print("   |------------------ 1.- CREAR INVENTARIO  --------------------------|")
    print("   |------------------ 2.- ENVIAR INVENTARIO --------------------------|")
    print("   |------------------ 3.-    SALIR          --------------------------|")
    print("   |-------------------------------------------------------------------|")
    
    option = int(input("\n Ingrese una opción: "))
    if(option == 1):
        headers()
        
    if(option == 3):
        break
