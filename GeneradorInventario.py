from Utils import headers
from Connection import connection
from openpyxl import load_workbook
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
        #Call file inventario.xlsx
        wb = load_workbook("inventario.xlsx")
        ws = wb.active
        #Conection Sql Server
        con = connection()
        sql = con.cursor()
        #Execute procedure
        command = "set nocount on; execute GeneradorDeInventarioXCategoriasV2 1,'2024-07-25' "
        sql.execute(command)
        rows = sql.fetchall()
        fila  = 5
        for row in rows:
          
            ws.cell(fila,1).value = row[0]
            ws.cell(fila,2).value = int(row[1])
            ws.cell(fila,3).value = row[2]
            ws.cell(fila,4).value = row[3]
            ws.cell(fila,5).value = row[4]
            ws.cell(fila,6).value = row[5]
            ws.cell(fila,11).value = row[6]
            ws.cell(fila,12).value = row[7]
            ws.cell(fila,13).value = row[8]
            ws.cell(fila,18).value = row[9]
            ws.cell(fila,19).value = row[10]
            ws.cell(fila,20).value = row[11]
           
            fila = fila+1
        #Add Total x category
        ws.cell(fila+1,3).value = "TOTALES"
        ws.cell(fila+1,6).value = "=SUM(F5:F"+str(fila-1)+")"
        ws.cell(fila+1,13).value = "=SUM(M5:M"+str(fila-1)+")"
        ws.cell(fila+1,20).value = "=SUM(T5:T"+str(fila-1)+")"
        wb.save("inventario.xlsx")  
        sql.close()
    if(option == 3):
        break
