from Utils import headers
from Connection import connection
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
        #Conection Sql Server
        con = connection()
        sql = con.cursor()
        #Execute procedure
        command = "set nocount on; execute GeneradorDeInventarioXCategoriasV2 1,'2024-07-24' "
        sql.execute(command)
        rows = sql.fetchall()
        for row in rows:
            print(row)
        sql.close()
    if(option == 3):
        break
