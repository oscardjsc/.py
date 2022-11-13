from ast import Continue
import datetime
import time

import openpyxl
import random

import sqlite3
from sqlite3 import Error

try:
    with sqlite3.connect("Reservaciones_Coworking.db") as conn:
        mi_cursor = conn.cursor()
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS cliente (num_cliente INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL);")
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS sala (num_sala INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL, cap_sala INTEGER NOT NULL);")
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS reservacion (folio_reservacion INTEGER PRIMARY KEY AUTOINCREMENT, num_sala INTEGER NOT NULL, num_cliente INTEGER NOT NULL, nombre_evento TEXT NOT NULL, fecha_reservacion timestamp, turno_reservacion TEXT NOT NULL,  FOREIGN KEY(num_sala) REFERENCES sala(num_sala), FOREIGN KEY(num_cliente) REFERENCES cliente(num_cliente));")
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS turno(folio INTEGER PRIMARY KEY AUTOINCREMENT, nombre_turno TEXT NOT NULL);")
        mi_cursor.execute("INSERT INTO turno VALUES(NULL,'Matutino');")
        mi_cursor.execute("INSERT INTO turno VALUES(NULL,'Vespertino');")
        mi_cursor.execute("INSERT INTO turno VALUES(NULL,'Nocturno');")
        print("Tablas creadas exitosamente")
finally:
  conn.close()

fecha_actual = datetime.date.today()
diferencia_dias = 2
fecha_reservacion_procesada = ""

lista_encontrados = []
reservaciones_posibles = []
reservaciones_posibles_final = []

while True:
    print("\n")
    print("Bienvenidos al sistema para la reservación de renta de espacios coworking")

    print("\t [A]Reservaciones")

    print("\t [B]Reportes")

    print("\t [C]Registrar una sala")

    print("\t [D]Registrar un cliente")

    print("\t [E]Salir")

    opcion=input("Elije la opción deseada, oprimiendo la tecla de la letra que corresponda: ")
    print("\n")
    print("*" * 60)

    if (not opcion.upper() in "ABCDE"):
            print("\n")
            print("OPCION INCORRECTA, POR FAVOR VUELVA A INTENTARLO")
            print("*" * 60)

    if (opcion.upper()== "A"):
        while True:
            print("\n")
            print("\t [A]Registrar una nueva reservación")
            print("\t [B]Modificar descripción de una reservación")
            print("\t [C]Consultar disponibilidad de una fecha")
            print("\t [D]Eliminar una reservación")
            print("\t [E]Volver al menú principal")

            opcion2=input("Elije la opción deseada, oprimiendo la tecla de la letra que corresponda: ")
            print("\n")
            print("*" * 60)

            if (not opcion2.upper() in "ABCDE"):
                print("OPCION INCORRECTA, POR FAVOR VUELVA A INTENTARLO")
                print("*" * 60)

            if (opcion2.upper()== "A"):
              while True:
                with sqlite3.connect("Reservaciones_Coworking.db") as conn:
                  mi_cursor = conn.cursor()
                  mi_cursor.execute("SELECT num_cliente, nombre FROM cliente")
                  registros = mi_cursor.fetchall()
                  print("\t CLIENTES")
                  print("*" * 60)
                  print(" NUMERO DE CLIENTE                NOMBRE")
                  for numero_cliente, nombre in registros:
                    print(f"          {numero_cliente}                   {nombre}")
                  print("*"*60)
                  print("\n")                
                try:
                  respuesta = int(input("Ingresar su número de cliente: "))
                  print("\n")
                  break
                except ValueError:
                  print("*" * 60)
                  print("El número de cliente no puede omitirse ni ser de caracter string, favor de intentarlo nuevamente")
                  print("*" * 60)
                  continue

              with sqlite3.connect("Reservaciones_Coworking.db") as conn:
                mi_cursor = conn.cursor()
                mi_cursor.execute("SELECT num_cliente FROM cliente")
                registros = mi_cursor.fetchall()
                if registros is None:
                  break
                  
                  
              if (respuesta,) in registros:
                while True:
                  with sqlite3.connect("Reservaciones_Coworking.db") as conn:
                    mi_cursor = conn.cursor()
                    mi_cursor.execute("SELECT num_sala, nombre, cap_sala FROM sala")
                    registros = mi_cursor.fetchall()
                    print("*" * 60)
                    print("\t SALAS")
                    print("*" * 60)
                    print(" SALA       NOMBRE       CAPACIDAD")
                    for sala, nombre, capacidad in registros:
                      print(f"   {sala}          {nombre}          {capacidad}")
                    print("*"*60)
                    print("\n")

                  try:
                    sala = int(input("Ingresa el numero de sala que sera utilizada: "))
                    print("*" * 60)
                    break
                  except ValueError:
                      print("*" * 60)
                      print("EL NUMERO DE SALA NO PUEDE OMITIRSE NI SER DE CARACTER ALFANUMERICO, FAVOR DE INTENTAR NUEVAMENTE")
                      print("*" * 60)
                      continue


                with sqlite3.connect("Reservaciones_Coworking.db") as conn:
                  mi_cursor = conn.cursor()
                  mi_cursor.execute("SELECT num_sala FROM sala")
                  registros = mi_cursor.fetchall()

                  if (sala,) in registros: 
                    while True:
                      nombre_evento=input("¿Cuál es el nombre de su evento?: ")
                      print("*" * 60)
                      if nombre_evento == "":
                        print("EL NOMBRE DEL EVENTO NO SE DEBE OMITIR, FAVOR DE INTENTARLO NUEVAMENTE")
                        print("*" * 60)
                      elif nombre_evento.isspace() is True:
                        print("EL NOMBRE DEL EVENTO NO PUEDEN SER ESPACIOS EN BLANCO, FAVOR DE INTENTARLO NUEVAMENTE")
                        print("*" * 60)
                      else:
                          break

                    while True:
                      try:
                        fecha_reservacion = input("¿Cuál sería la fecha en que deseea realizar su reservación? (DD/MM/AAAA): ")
                        print("*" * 60)
                        if fecha_reservacion == "":
                          print("LA FECHA DE RESERVACION NO SE DEBE OMITIR, FAVOR DE INTENTARLO NUEVAMENTE")
                          print("*" * 60)
                          continue
                        fecha_reservacion_procesada = datetime.datetime.strptime(fecha_reservacion, "%d/%m/%Y").date()
                        break
                      except ValueError:
                        print("FORMATO DE FECHA NO VALIDO, FAVOR DE INTENTARLO NUEVAMENTE")
                        print("*" * 60)

                    diferencia_dias=fecha_reservacion_procesada - fecha_actual
                    if diferencia_dias.days <=1:
                      print("LA RESERVACION DE UNA SALA, TIENE QUE SER POR LO MINIMO CON 2 DIAS DE ANTERIORIDAD")
                      print("*" * 60)
                      break
                    else:
                      while True:
                        try:
                          turno_reservacion = int(input("Favor de seleccionar el turno deseado (1. Matutino, 2. Vespertino, 3. Nocturno): "))
                          print("*" * 60)
                          if turno_reservacion == 1:
                            turno_reservacion = "Matutino"
                            break
                          elif turno_reservacion == 2:
                             turno_reservacion = "Vespertino"
                             break 
                          elif turno_reservacion == 3:
                             turno_reservacion = "Nocturno"
                             break    
                          else:
                            print("ELECCION ERRONEA, FAVOR DE INTENTARLO NUEVAMENTE")
                            print("*" * 60)
                            continue
                        except ValueError:
                            print("EL NUMERO DE SALA NO PUEDE OMITIRSE NI SER DE CARACTER ALFANUMERICO, FAVOR DE INTENTARLO NUEVAMENTE")
                            print("*" * 60)
                            continue
                           
                            

                    with sqlite3.connect("Reservaciones_Coworking.db", detect_types = sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
                      mi_cursor = conn.cursor()
                      valores = {"num_sala":sala, "nombre_evento":nombre_evento, "fecha_reservacion":fecha_reservacion_procesada, "turno_reservacion":turno_reservacion}
                      mi_cursor.execute("SELECT count(*) FROM reservacion WHERE num_sala=:num_sala AND DATE(fecha_reservacion)=:fecha_reservacion AND turno_reservacion=:turno_reservacion", valores)
                      registros2 = mi_cursor.fetchall()
                                            

                      if (1,) in registros2:
                        print("LO SENTIMOS, YA EXISTE UNA RESERVACION PARA ESTA SALA, EN EL DIA Y TURNO SELECCIONADO")
                        print("*" * 60)
                      else:
                        with sqlite3.connect("Reservaciones_Coworking.db", detect_types = sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
                          mi_cursor = conn.cursor()
                          valores = {"num_sala":sala,"cliente":respuesta, "nombre_evento":nombre_evento, "fecha_reservacion":fecha_reservacion_procesada, "turno_reservacion":turno_reservacion}
                          mi_cursor.execute("INSERT INTO reservacion VALUES(NULL, :num_sala,:cliente, :nombre_evento, :fecha_reservacion, :turno_reservacion)", valores)
                        with sqlite3.connect("Reservaciones_Coworking.db") as conn:
                          mi_cursor = conn.cursor()
                          valores = {"num_sala":sala, "nombre_evento":nombre_evento, "fecha_reservacion":fecha_reservacion_procesada, "turno_reservacion":turno_reservacion}
                          mi_cursor.execute("SELECT folio_reservacion FROM reservacion WHERE num_sala=:num_sala AND DATE(fecha_reservacion)=:fecha_reservacion AND turno_reservacion=:turno_reservacion", valores)
                          registros = mi_cursor.fetchall()  
                          for folio, in registros:
                            folio_imprimir=folio                                             
                          print(f"REGISTRO AGREGADO EXITOSAMENTE, SU NUMERO DE FOLIO DE RESERVACION ES:  {folio_imprimir} ")
                          print("\n")
                  else:
                      print("SALA SELECCIONADA NO EXISTENTE")
                      print("*" * 60)
                      print("\n")
                      
              else:
                  print("PARA REALIZAR UNA RESERVACION ES NECESARIO SER CLIENTE REGISTRADO, FAVOR DE PRIMERO HACER SU REGISTRO ")
                  print("*" * 60)
                  print("\n")
                  


            if (opcion2.upper()== "B"):
              while True:
                with sqlite3.connect("Reservaciones_Coworking.db") as conn:
                  mi_cursor = conn.cursor()
                  mi_cursor.execute("SELECT folio_reservacion, num_sala, num_cliente, nombre_evento, fecha_reservacion, turno_reservacion  FROM reservacion")
                  registros = mi_cursor.fetchall()
                  print("\t RESERVACIONES")
                  print("*" * 130)
                  print(" FOLIO       NUMERO DE SALA       NUMERO DE CLIENTE        NOMBRE DE EVENTO      FECHA DE RESERVACION      TURNO DE RESERVACION")
                  for folio, sala, cliente, nombre, reservacion, turno in registros:
                    print(f"   {folio}                 {sala}                      {cliente}                   {nombre}              {reservacion}               {turno} ")
                  print("*"*130)
                  print("\n")
                try:
                    folio_reservacion = int(input("¿Cuál es el número de folio de su reservación?: "))
                    print("*" * 60)
                    break
                except ValueError:
                  print("El folio de reservación no se puede omitir ni ser caracteres alfanuméricos, favor de intentarlo nuevamente")
                  print("*" * 60)


              with sqlite3.connect("Reservaciones_Coworking.db") as conn:
                mi_cursor = conn.cursor()
                valores = {"folio_reservacion":folio_reservacion}
                mi_cursor.execute("SELECT folio_reservacion FROM reservacion WHERE folio_reservacion=:folio_reservacion", valores)
                registros = mi_cursor.fetchall()


                if (folio_reservacion,) in registros:
                  while True:
                    nuevo_nombre = input("¿Cuál será el nuevo nombre de su evento: ")
                    if nuevo_nombre == "":
                      print("EL NUEVO NOMBRE NO SE DEBE DE OMITIR, FAVOR DE INTENTAR NUEVAMENTE")
                      print("*" * 60)
                    elif nuevo_nombre.isspace() is True:
                      print("EL NUEVO NOMBRE NO PUEDE SER ESPACIOS EN BLANCO, FAVOR DE INTENTAR NUEVAMENTE")
                      print("*" * 60)
                    else:
                      break                                      
                      print("*" * 60)

                  with sqlite3.connect("Reservaciones_Coworking.db") as conn:
                    mi_cursor = conn.cursor()
                    valores = {"nuevo_nombre":nuevo_nombre, "folio_reservacion":folio_reservacion} 
                    mi_cursor.execute("UPDATE reservacion SET nombre_evento=:nuevo_nombre WHERE folio_reservacion=:folio_reservacion", valores)
                    print("*" * 60)
                    print("DESCRIPCION MODIFICADA EXITOSAMENTE")
                    print("*" * 60)
                else:
                    print("*" * 60)
                    print("EL NUMERO DE FOLIO DE RESERVACION NO FUE ENCONTRADO")
                    print("*" * 60)



                
            if (opcion2.upper()== "C"):
              while True:
                try:
                  fecha_buscada=input("Ingresa la fecha buscada:")
                  print("\n")
                  if fecha_buscada == "":
                    print("LA FECHA BUSCADA NO SE DEBE OMITIR, FAVOR DE INTENTAR NUEVAMENTE")
                    print("*" * 60)
                    continue
                  fecha_reservacion_procesada2 = datetime.datetime.strptime(fecha_buscada, "%d/%m/%Y").date() 
                  break
                except ValueError:
                  print("FORMATO DE FECHA NO VALIDO, FAVOR DE INTENTAR NUEVAMENTE")
                  print("*" * 60)    

              with sqlite3.connect("Reservaciones_Coworking.db", detect_types = sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
                mi_cursor = conn.cursor()
                valores = {"fecha_buscada":fecha_reservacion_procesada2}
                mi_cursor.execute("SELECT R.num_sala, S.nombre, R.turno_reservacion FROM reservacion R INNER JOIN sala S ON R.num_sala=S.num_sala WHERE DATE(R.fecha_reservacion)=:fecha_buscada", valores)
                registros = mi_cursor.fetchall() 
                  
                lista_encontrados.clear()
                for sala, nombre, turno in registros:
                  lista_encontrados.append((sala,nombre,turno))
                
                reservaciones_encontradas=set(lista_encontrados)    
                              

                mi_cursor.execute("SELECT num_sala, nombre FROM sala")
                registros2 = mi_cursor.fetchall() 
                mi_cursor.execute("SELECT nombre_turno FROM turno WHERE folio IN (1,2,3)")
                registros3 = mi_cursor.fetchall()

                for sala, nombre in registros2:
                  for eleccion, in registros3:
                    reservaciones_posibles.append((sala,nombre,eleccion))

                for elemento in reservaciones_posibles:
                  if elemento not in reservaciones_posibles_final:
                      reservaciones_posibles_final.append(elemento)
                
                reservaciones_posibles_realizar=set(reservaciones_posibles_final)
                total=  sorted(reservaciones_posibles_realizar - reservaciones_encontradas)
                print("*" * 60)
                print(f'LA DISPONIBILIDAD PARA EL DIA {fecha_buscada} ES LA SIGUIENTE: ')
                print("*" * 60)
                print("Sala        Turno")
                for datos in total:
                  print(f'{datos[0]},{datos[1]}     {datos[2]}     ')
                print("*" * 60)
               


            if (opcion2.upper()== "D"):
              while True:
                with sqlite3.connect("Reservaciones_Coworking.db") as conn:
                  mi_cursor = conn.cursor()
                  mi_cursor.execute("SELECT folio_reservacion, num_sala, num_cliente, nombre_evento, fecha_reservacion, turno_reservacion  FROM reservacion")
                  registros = mi_cursor.fetchall()
                  print("\t RESERVACIONES")
                  print("*" * 130)
                  print(" FOLIO       NUMERO DE SALA       NUMERO DE CLIENTE        NOMBRE DE EVENTO      FECHA DE RESERVACION      TURNO DE RESERVACION")
                  for folio, sala, cliente, nombre, reservacion, turno in registros:
                    print(f"   {folio}                 {sala}                      {cliente}                   {nombre}              {reservacion}               {turno} ")
                  print("*"*130)
                  print("\n")
                try:
                  folio =int(input("Ingresa el numero de folio de la reservacion que deseas eliminar: "))
                  print("*" * 60)
                  break
                except ValueError:
                  print("EL FOLIO DE LA RESERVACION A ELIMINAR, NO SE DEBE DE OMITIR NI SER CARACTERES ALFANUMERICOS, FAVOR DE INTENTARLO NUEVAMENTE")
                  print("*" * 60)
              while True:
                try:
                  fecha_eliminar = input("Ingresa la fecha en la que la reservacion esta agendada (DD/MM/AAAA): ")
                  print("\n")
                  if fecha_eliminar == "":
                    print("LA FECHA BUSCADA NO SE DEBE DE OMITIR, FAVOR DE INTENTARLO NUEVAMENTE")
                    print("*" * 60)
                    continue
                  fecha_eliminar_procesada = datetime.datetime.strptime(fecha_eliminar, "%d/%m/%Y").date()
                  break
                except ValueError:
                  print("FORMATO DE FECHA NO VALIDO, FAVOR DE VOLVER A INTENTARLO")
                  print("*" * 60)
                  
              with sqlite3.connect("Reservaciones_Coworking.db", detect_types = sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
                mi_cursor = conn.cursor()
                valores = {"folio":folio, "fecha_eliminar":fecha_eliminar_procesada}
                mi_cursor.execute("SELECT count(*) FROM reservacion WHERE folio_reservacion=:folio AND fecha_reservacion=:fecha_eliminar ", valores)
                registros = mi_cursor.fetchall()
                                        
                if (0,) in registros:
                  print("NO SE ENCONTRO NINGUN FOLIO CON EL NUMERO Y FECHA INGRESADA, FAVOR DE VERIFICAR LOS VALORES INGRESADOS")
                  print("*" * 60)
                  break
                else:
                   with sqlite3.connect("Reservaciones_Coworking.db", detect_types = sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
                      mi_cursor = conn.cursor()
                      valores = {"folio":folio, "fecha_eliminar":fecha_eliminar_procesada}
                      mi_cursor.execute("SELECT folio_reservacion, num_sala, num_cliente, nombre_evento, DATE(fecha_reservacion), turno_reservacion FROM reservacion WHERE folio_reservacion=:folio AND DATE(fecha_reservacion)=:fecha_eliminar", valores)
                      registros = mi_cursor.fetchall()

                      print("LOS DATOS DEL FOLIO DE RESERVACION INGRESADOS SON LOS SIGUIENTES:")
                      print("*" * 60)
                      for folio_reservacion, num_sala, num_cliente, nombre_evento, fecha_reservacion, turno_reservacion  in registros:
                          print(f"Folio de Reservacion = {folio_reservacion}")
                          print(f"Numero de Sala = {num_sala}")
                          print(f"Numero de Cliente = {num_cliente}")
                          print(f"Nombre de Evento = {nombre_evento}")
                          print(f"Fecha de reservacion = {fecha_reservacion}")
                          print(f"Turno de reservacion = {turno_reservacion}")

                      
                      diferencia_dias2=fecha_eliminar_procesada - fecha_actual
                      if diferencia_dias2.days <=2:
                        print("*" * 60)
                        print("La eliminacion de la reservacion no se puede realizar a menos de 3 dias de anticipacion del evento")
                        break
                      else:
                        print("*" * 60)
                        print("¡RECUERDA QUE AL ELIMINAR UNA RESERVACION, ESO NO PODRA SER DESECHO!")
                        print("\n")
                        while True:
                          try:
                            decision_eliminar=int(input("Desea eliminar la reservacion: 1.Si  2.No: "))
                            print("*" * 60)
                            break
                          except ValueError:
                            print("LA RESPUESTA INGRESADA NO FUE CORRECTA, FAVOR DE VOLVER A INTENTARLO")
                        print("*" * 60)
                        if decision_eliminar==1:
                          with sqlite3.connect("Reservaciones_Coworking.db", detect_types = sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
                              mi_cursor = conn.cursor()
                              valores = {"folio":folio, "fecha_eliminar":fecha_eliminar_procesada}
                              mi_cursor.execute("DELETE FROM reservacion WHERE folio_reservacion=:folio", valores)
                              registros = mi_cursor.fetchall()
                              print("Eliminacion de la reservacion hecha con exito")
                              print("*" * 60)
                        elif decision_eliminar==2:
                          print("Eliminacion cancelada")
                          print("*" * 60)
                        else:
                          print("OPCION INCORRECTA, ELIMINACION NO REALIZADA")



            if (opcion2.upper()== "E"):
              break

    if (opcion.upper()== "B"):
        while True:
            print("\t [A]Reporte en pantalla de reservaciones en una fecha")
            print("\t [B]Exportar reporte tabular en excel")
            print("\t [C]Volver al menu principal")

            opcion3=input("Elije la opcion deseada, oprimiendo la tecla de la letra que corresponda: ")
            print("\n")
            print("*" * 60)

            if (not opcion3.upper() in "ABC"):
                print("OPCION INCORRECTA, FAVOR DE VOLVER A INTENTARLO")
                print("*" * 60)

            if (opcion3.upper()=="A"):
              while True:
                try:
                  fecha_mostrar= input("¿Cuál sería la fecha en que deseea ver las reservaciones realizadas? (DD/MM/AAAA): ")
                  print("\n")
                  if fecha_mostrar == "":
                    print("NO SE INGRESO NINGUN DATO, FAVOR DE VOLVER A INTENTARLO")
                    print("*"*60)
                    continue
                  fecha_mostrar_procesada = datetime.datetime.strptime(fecha_mostrar, "%d/%m/%Y").date()
                  break
                except ValueError:
                  print("FORMATO DE FECHA NO VALIDO, FAVOR DE VOLVER A INTENTARLO")
                  print("*"*60)

              with sqlite3.connect("Reservaciones_Coworking.db", detect_types = sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
                mi_cursor = conn.cursor()
                valores = {"fecha_mostrar":fecha_mostrar_procesada}
                mi_cursor.execute("SELECT num_sala, num_cliente, nombre_evento, turno_reservacion FROM reservacion WHERE DATE(fecha_reservacion)=:fecha_mostrar", valores)
                registros = mi_cursor.fetchall() 

                if registros:
                  print("*" * 60)
                  print(f'\t REPORTE DE RESERVACIONES DEL DIA {fecha_mostrar}')
                  print("*" * 60)
                  print("Sala    Numero de Cliente       Nombre de reservacion        Turno")
                  print("*" * 80)
                  for sala, cliente, evento, turno in registros:
                    print(f'{sala}                {cliente}                     {evento}              {turno}')
                    print("\n")
                else:
                  print("No se cuentan con reservaciones agendadas para la fecha buscada")
                  print("*" * 60)
                  print("\n")
            
            if (opcion3.upper()=="B"):
              while True:
                try:
                  fecha_exportar= input("¿Cuál sería la fecha en que deseea obtener el reporte de reservaciones en Excel? (DD/MM/AAAA):")
                  print("\n")
                  if fecha_exportar == "":
                    print("NO SE INGRESO NINGUN DATO, FAVOR DE VOLVER A INTENTARLO")
                    print("*" * 60)
                    continue
                  fecha_exportar_procesada = datetime.datetime.strptime(fecha_exportar, "%d/%m/%Y").date()
                  break
                except ValueError:
                  print("FORMATO DE FECHA NO VALIDO, FAVOR DE VOLVER A INTENTARLO")
                  print("*" * 60)

              with sqlite3.connect("Reservaciones_Coworking.db", detect_types = sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
                mi_cursor = conn.cursor()
                valores = {"fecha_exportar":fecha_exportar_procesada}
                mi_cursor.execute("SELECT R.folio_reservacion, S.nombre, R.num_cliente, R.nombre_evento, DATE(R.fecha_reservacion), R.turno_reservacion FROM reservacion R INNER JOIN sala S ON R.num_sala=S.num_sala WHERE DATE(fecha_reservacion)=:fecha_exportar", valores)
                registros = mi_cursor.fetchall()  

                if registros:
                  wb = openpyxl.Workbook()
                  hoja1 = wb.create_sheet("Hoja")
                  hoja = wb.active
                  hoja.append(('Folio de reservacion','Nombre de Sala', ' Numero del cliente', 'Nombre del evento', 'Fecha', 'Turno'))

                  for reservaciones in registros:
                    hoja.append(reservaciones)
                    wb.save(f'Reservaciones_{fecha_exportar_procesada}.xlsx')
                  print("Datos exportados correctamente a un archivo de MsExcel")
                  print("*" * 60)
                  print("\n")
                else:
                  print("No se cuentan con reservaciones agendadas para ese dia, por lo tanto no hubo exportación de datos")
                  print("*" * 60)
                  print("\n")

                  

            if (opcion3.upper()=="C"):
                break


    if (opcion.upper()== "C"):
        while True:
          nombre_sala=input("Ingresa el nombre de la sala: ")
          print("*" * 60)
          if nombre_sala == "":
            print("EL NOMBRE DE LA SALA NO DEBE DE OMITIRSE, INTENTELO NUEVAMENTE")
            print("*" * 60)
          elif nombre_sala.isspace() is True:
            print("*" * 60)
            print("EL NOMBRE DE LA SALA NO PUEDE SER ESPACIOS EN BLANCO, FAVOR DE VOLVER A INTENTARLO")
            print("*" * 60)
          else:
            break
        while True:
          try:
            cap_sala=int(input("Ingresa la cantidad de aforo maximo de la sala: "))
            print("*"*60)
            if cap_sala <=0:
              print("EL CUPO NO PUEDE OMITIRSE Y/O SER MENOR A 0, INTENTELO NUEVAMENTE ")
              print("*" * 60)
              continue
            else:
              print("*" * 60)
              break
          except ValueError:
            print("*" * 60)
            print("EL NUMERO DE SALA NO PUEDE OMITIRSE NI SER DE CARACTER ALFANUMERICO, FAVOR DE INTENTARLO NUEVAMENTE")
            print("*" * 60)
            continue
        try:
          with sqlite3.connect("Reservaciones_Coworking.db") as conn:
              mi_cursor = conn.cursor()
              valores = {"nombre_sala":nombre_sala, "cap_sala":cap_sala}
              mi_cursor.execute("INSERT INTO sala VALUES(NULL, :nombre_sala, :cap_sala)", valores)
              print("REGISTRO AGREGADO EXITOSAMENTE")
        finally:
            conn.close()
        try:
            with sqlite3.connect("Reservaciones_Coworking.db") as conn:
              mi_cursor = conn.cursor()
              mi_cursor.execute("SELECT num_sala FROM sala WHERE nombre=:nombre_sala AND cap_sala=:cap_sala", valores)
              registros = mi_cursor.fetchall()

              for sala, in registros:
                  sala_imprimir=sala
              print(f'El numero de la sala asignado es: {sala_imprimir}')
        finally:
            conn.close()        
       

        
    if (opcion.upper()== "D"):
        while True:
          nombre_cliente=input("Ingrese su nombre completo: ")
          print("*" * 60)
          if nombre_cliente == "":
            print("EL NOMBRE DEL CLIENTE NO SE DEBE DE OMITIR, FAVOR DE INTENTARLO NUEVAMENTE")
            print("*" * 60)
          elif nombre_cliente.isspace() is True:
            print("EL NOMBRE DEL CLIENTE NO PUEDEN SER ESPACIOS EN BLANCO, FAVOR DE VOLVER A INTENTARLO")
            print("*" * 60)
          else:
            break
        try:
          with sqlite3.connect("Reservaciones_Coworking.db") as conn:
              mi_cursor = conn.cursor()
              valores = {"nombre":nombre_cliente}
              mi_cursor.execute("INSERT INTO cliente VALUES(NULL, :nombre)", valores)
              print("Registro agregado exitosamente")
        finally:
            conn.close()
        try:
            with sqlite3.connect("Reservaciones_Coworking.db") as conn:
              mi_cursor = conn.cursor()
              mi_cursor.execute("SELECT num_cliente FROM cliente WHERE nombre=:nombre", valores)
              registros = mi_cursor.fetchall()

              for cliente, in registros:
                cliente_imprimir=cliente
              print(f'Su numero de cliente asignado es: {cliente_imprimir}')
        finally:
            conn.close()


    if (opcion.upper()== "E"):
      print("¡QUE TENGA UN BONITO DÍA!")
      break
