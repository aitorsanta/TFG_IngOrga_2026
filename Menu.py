# ============================================================
# AUTORÍA DEL CÓDIGO
# Autor: Aitor Santamaria Zuluaga
# Grado en Ingeniería en Organización Industrial
# JULIO DEL 2026
# ============================================================

import Asignacion
import horarioEscolar as Horarios

def menu_principal():
    # Asegura que el archivo base existe al arrancar
    Asignacion.crear_excel_base()
    
    while True:
        print("\n" + "="*46)
        print("      GESTOR ESCOLAR TFG AITOR SANTAMARIA")
        print("="*46)
        print("1. Asignar Profesores (Generar asignacion_completa.xlsx)")
        print("2. Ver Datos Actuales (Asignaturas y Profesores)")
        print("3. Generar/Resetear Hoja de Afinidades")
        print("4. GENERAR HORARIOS")
        print("5. Salir")
        print("="*46)

        opcion = input("Elige una opción: ")

        if opcion == "1":
            print("\n[1/4] Cargando configuraciones...")
            asignaturas = Asignacion.cargar_asignaturas_excel()
            profesores = Asignacion.cargar_profesores_excel()
            
            # Cargamos los desdobles desde Asignacion para validar la carga de horas
            desdobles = Asignacion.cargar_desdobles_excel()

            if not asignaturas or not profesores:
                print("[!] Error: Faltan datos base en datos_centro.xlsx")
                continue

            print("[2/4] Cargando matriz de afinidades...")
            afinidades = Asignacion.cargar_afinidades_excel(profesores, asignaturas)

            if not afinidades:
                # Si no hay afinidades, el método cargar_afinidades ya avisa de que se ha creado la hoja
                continue

            print("[3/4] Ejecutando algoritmo de asignación inteligente...")
            resultado = Asignacion.asignar_profesores_automaticamente(
                profesores,
                asignaturas,
                afinidades,
                desdobles
            )

            if resultado is not None:
                asignacion, carga_profesor = resultado
                print("[4/4] Exportando resultados...")
                Asignacion.imprimir_carga_horaria(profesores, carga_profesor)
                Asignacion.exportar_asignacion_excel(
                    profesores,
                    asignaturas,
                    asignacion,
                    carga_profesor,
                    Asignacion.ARCHIVO_EXCEL_SALIDA
                )
                print(f"\n[OK] Asignación completada y guardada en {Asignacion.ARCHIVO_EXCEL_SALIDA}")
            else:
                print("\n[!] La asignación automática ha fallado. Revisa las cargas horarias.")

        elif opcion == "2":
            asigs = Asignacion.cargar_asignaturas_excel()
            profes = Asignacion.cargar_profesores_excel()
            
            print(f"\n--- {len(asigs)} Asignaturas ---")
            for a in asigs:
                bloque = f" [Bloque: {a.get('BloqueDesdoble')}]" if a.get('BloqueDesdoble') else ""
                print(f"- {a['Asignatura']} ({a['Grupo']}): {a['Horas']}h{bloque}")

            print(f"\n--- {len(profes)} Profesores ---")
            for p in profes:
                print(f"- {p['Nombre']} {p['Apellido1']} (Max: {p['MaxHoras']}h)")

        elif opcion == "3":
            asignaturas = Asignacion.cargar_asignaturas_excel()
            profesores = Asignacion.cargar_profesores_excel()
            if asignaturas and profesores:
                Asignacion.generar_hoja_afinidades_excel(profesores, asignaturas)
                print("\n[OK] Hoja 'Afinidades' generada en datos_centro.xlsx. Rellénala con 0, 1 o 2.")
            else:
                print("[!] No hay datos suficientes para generar afinidades.")

        elif opcion == "4":
            print("\n" + "*"*40)
            print("  INICIANDO MOTOR DE OPTIMIZACIÓN CP-SAT")
            print("*"*40)
            # El motor leerá directamente de datos_centro y asignacion_completa
            Horarios.generar_horarios()
            input("\nPresiona Intro para continuar...")

        elif opcion == "5":
            print("Cerrando el sistema.")
            break
        else:
            print("Opción no válida.")

if __name__ == "__main__":
    menu_principal()