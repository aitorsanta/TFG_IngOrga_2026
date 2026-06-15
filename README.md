# TFG Ingeniería en Organización Industrial 2026 · AITOR

**Datos de la aplicación:**
**Título:** Optimización y generación automática de horarios escolares mediante programación por restricciones: implementación de una aplicación en Python.\
**Autor:** Aitor Santamaria Zuluaga.\
Grado en Ingeniería en Organización Industrial. Universidad Internacional de la Rioja (UNIR).\
Curso 2025-2026.\
Licencia: Proyecto desarrollado con fines académicos como Trabajo Fin de Grado.

Descripción
Este proyecto ha sido desarrollado como Trabajo Fin de Grado del Grado en Ingeniería en Organización Industrial. El objetivo principal es automatizar y optimizar dos de los problemas organizativos más complejos de un centro educativo: la asignación de profesorado a materias y grupos y la generación automática de horarios escolares.\
Para ello se ha diseñado una herramienta completa en Python capaz de leer la estructura académica de un centro educativo desde hojas Excel, asignar docentes teniendo en cuenta afinidades y restricciones, y generar horarios válidos mediante técnicas de optimización basadas en programación por restricciones. El sistema ha sido desarrollado pensando en centros de Educación Secundaria Obligatoria y Bachillerato, aunque su estructura permite adaptaciones a otros contextos educativos.

**Funcionalidades principales**
1) Asignación automática de profesorado:
- Definir materias, grupos y carga horaria.
- Definir profesorado y carga máxima semanal.
- Establecer afinidades entre docentes y materias.
- Gestionar asignaciones obligatorias.
- Gestionar desdobles y agrupamientos.
- Obtener una propuesta automática de reparto docente.

2) Generación automática de horarios
- Evitar solapamientos de profesorado.
- Evitar conflictos de aulas.
- Gestionar aulas específicas y aulas flexibles.
- Respetar disponibilidades del profesorado.
- Gestionar reuniones de equipos docentes.
- Gestionar restricciones temporales.
- Controlar máximos diarios por asignatura.
- Gestionar bloques de horas consecutivas.
- Generar horarios completos para grupos y profesores.

3) Exportación de resultados
La aplicación genera automáticamente archivos Excel con:
- Asignación docente.
- Horarios de grupos.
- Horarios individuales de profesorado.
- Resúmenes de tutorías.
- Resúmenes de ocupación de aulas.
- Relación profesor-grupo.

**Tecnologías utilizadas**
Python
OR-Tools (Google CP-SAT Solver)
Pandas
OpenPyXL

**Módulos principales**
1) Menu.py: punto de entrada de la aplicación. Gestiona el menú principal y coordina la ejecución de los distintos módulos.
2) Asignacion.py: implementa la asignación automática de profesorado a materias y grupos.
3) horarioEscolar.py: contiene el modelo de optimización basado en programación por restricciones y la generación de horarios.
4) Horario.py: configuración inicial de la plantilla.

**Ejecución**
Completar los datos del centro en datos_centro.xlsx.
Ejecutar:
1) python Menu.py
2) Seleccionar la opción deseada del menú.
3) Generar la asignación docente.
4) Generar los horarios.
5) Los resultados se exportarán automáticamente a archivos Excel.
