# Script para Comparar Excel Base y Levantamiento en Campo

Este script en MATLAB permite comparar los datos provenientes de un archivo Excel base y los datos obtenidos en el levantamiento en campo. Se encarga de:
- Leer datos de distintas hojas de Excel.
- Convertir coordenadas geográficas a UTM.
- Calcular la distancia entre puntos de la base y el campo.
- Generar alertas en función de la discrepancia de las ubicaciones.

---

## Información del Proyecto

- **Desarrollador:** Juan Ortiz (2024)
- **Empresa:** CONSORCIO SIG-ELECTRIC
- **Versión:** 1

---

## Descripción

El script realiza las siguientes operaciones:

1. **Inicialización y Configuración:**  
   - Limpia el entorno de trabajo (`clc`, `clear`) y establece el formato de salida.
   - Añade las rutas necesarias donde se encuentran las funciones auxiliares y los archivos de Excel.

2. **Ingreso de Datos:**  
   - Se definen las variables del alimentador y las rutas de los archivos de Excel de entrada y salida.
   - Se especifican las hojas a leer en el Excel base y en el archivo de levantamiento en campo.

3. **Lectura de Datos:**  
   - Utiliza la función `F_leerDatosExcel` para importar los datos de las hojas especificadas en los archivos de Excel.

4. **Procesamiento y Cálculo de Distancias:**  
   - Extrae las coordenadas geográficas (latitud y longitud) de los medidores del levantamiento en campo.
   - Convierte estas coordenadas a UTM utilizando la función `deg2utm_ecuador`.
   - Calcula la distancia euclidiana entre las coordenadas de la base y las coordenadas obtenidas en campo.
   - Genera alertas:
     - **"A PC Desplazado"** si la distancia es mayor a 10 metros.
     - **"B Revisar Ubicacion"** si la distancia supera los 100 metros.

5. **Exportación de Resultados:**  
   - Escribe los datos procesados y las alertas generadas en hojas de Excel para su posterior análisis.

6. **Funciones Auxiliares:**  
   - La función `deg2utm_ecuador` se encarga de convertir las coordenadas geográficas (lat, lon) a coordenadas UTM, considerando el hemisferio.

---

## Estructura del Script

- **Inicialización:**  
  Limpieza de variables y configuración del entorno.

- **Ingreso de Datos:**  
  Definición de rutas y variables que contienen la información de entrada y salida.

- **Lectura de Datos:**  
  Importación de datos desde Excel mediante la función `F_leerDatosExcel`.

- **Procesamiento de Datos:**  
  Conversión de coordenadas, cálculo de distancias y generación de alertas.

- **Exportación de Resultados:**  
  Escritura de los resultados y alertas en archivos Excel.

- **Función `deg2utm_ecuador`:**  
  Conversión de coordenadas geográficas a UTM.

---

## Requisitos

- **MATLAB R2019b o superior:**  
  Se recomienda utilizar una versión reciente de MATLAB para asegurar la compatibilidad con las funciones utilizadas.

- **Funciones Auxiliares:**  
  Asegúrate de que la función `F_leerDatosExcel` y cualquier otro archivo en la carpeta `F_Funciones\` estén correctamente ubicados en el repositorio.

- **Acceso a Archivos:**  
  Verifica que las rutas definidas para los archivos de Excel de entrada y salida sean correctas y estén accesibles desde tu entorno de trabajo.

---

## Uso

1. **Clonar el Repositorio:**  
   Clona este repositorio en tu máquina local.

2. **Configurar Rutas y Variables:**  
   Modifica las rutas de los archivos de entrada y salida según tu estructura de directorios (en las variables `inputFileBase`, `folderOut`, etc.).

3. **Ejecutar el Script:**  
   Abre MATLAB, navega hasta la carpeta del proyecto y ejecuta el script.

4. **Revisar Resultados:**  
   Los archivos Excel generados con los resultados y las alertas se guardarán en la carpeta especificada para la salida.

---

## Contribuciones

Si deseas mejorar o agregar funcionalidades al script, siéntete libre de hacer un *fork* y enviar un *pull request*. Toda contribución es bienvenida.

---

## Licencia

Este proyecto es propiedad de **CONSORCIO SIG-ELECTRIC**. Consulta el archivo `LICENSE` para conocer los términos y condiciones de uso.

---

Para cualquier consulta o comentario, por favor contacta al desarrollador **Juan Ortiz**.

