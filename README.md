Autores: Juan Herrera, Juan Castiblanco, Andres Moncada y Mauricio Erazo.
# 📊 Limpieza y Formateo Automático de Datos de Cultivos en Excel

Este script de Python está diseñado para automatizar dos tareas clave en un archivo de Excel:
1. **Limpiar** el conjunto de datos eliminando filas que contienen la etiqueta 'TOTAL' en la columna 'Cultivo'.
2. **Aplicar un formato estético profesional** a los datos restantes (encabezados, bordes, ancho de columna) y ajustar el encabezado 'Valor' a 'Valor (%)'.

## ⚙️ Flujo de Trabajo del Algoritmo (Workflow)

El algoritmo se ejecuta en tres fases principales dentro del bloque `if __name__ == '__main__':`.

---

### Fase 1: Configuración y Rutas 🗺️

1.  **Definición de Rutas:** Se establecen las rutas de entrada (`INPUT_PATH`) y de salida (`OUTPUT_PATH`) del archivo Excel.
    * **Entrada:** `Partefinal1.xlsx` (o el nombre configurado).
    * **Salida:** `Finalreto1.xlsx` (o el nombre configurado) dentro del directorio de salida.
2.  **Montaje de Entorno (Opcional):** Se incluye un *placeholder* para el montaje de Google Drive (`from google.colab import drive`), ya que el script está diseñado para ejecutarse en entornos como Google Colab.

---

### Fase 2: Limpieza y Filtrado de Datos 🧹

Esta fase está a cargo de la función `limpiar_eliminar_total(input_path, output_path, total_label)`.

1.  **Lectura de Datos:** Utiliza **Pandas** (`pd.read_excel`) para leer el archivo de Excel especificado en `INPUT_PATH` en un `DataFrame`.
2.  **Filtrado (Paso Crítico):** Se filtra el `DataFrame` para **excluir** todas las filas donde el valor de la columna `'Cultivo'`, convertido a mayúsculas, coincide *exactamente* con la etiqueta `TOTAL` (o la etiqueta definida por `total_label`).
    * `df_limpio = df[df['Cultivo'].str.upper() != total_label].copy()`
3.  **Guardado Temporal:** El `DataFrame` limpio se guarda inmediatamente en el disco con el nombre del archivo de salida (`OUTPUT_PATH`). Esto es crucial porque la siguiente fase (`aplicar_estetica_y_guardar`) utiliza la librería `openpyxl`, que requiere un archivo guardado para cargar el libro de trabajo.

---

### Fase 3: Aplicación de Estética y Formato 🎨

Esta fase está a cargo de la función `aplicar_estetica_y_guardar(df_limpio, output_path)` y aplica el formateo utilizando **OpenPyXL**.

1.  **Carga del Archivo:** El archivo Excel recién guardado en `OUTPUT_PATH` se carga en un objeto `Workbook` de `openpyxl`.
2.  **Detección de Columna 'Valor':** Se busca la posición de la columna `'Valor'` en el `DataFrame` limpio para poder modificar su encabezado y aplicar formatos si fuera necesario.

#### A. Formato de Encabezados (Fila 1)
* Se itera sobre todas las celdas de la primera fila.
* A cada celda de encabezado se le aplica el siguiente estilo:
    * **Fuente:** Negrita (`FUENTE_NEGRITA`).
    * **Relleno:** Color azul (`AZUL_RELLENO`).
    * **Alineación:** Centro (`ALINEACION_CENTRO`).
    * **Bordes:** Borde fino en todos los lados (`BORDE_FINO`).

#### B. Formato y Ajuste de Datos
* Se itera sobre **cada celda** de la hoja de cálculo (incluyendo encabezados y datos).
* **Bordes:** A todas las celdas se les aplica un borde fino.
* **Alineación de Datos:** A todas las celdas **excepto** la fila de encabezados se les aplica alineación centrada.
* **Ancho de Columna:** Se calcula la longitud máxima del contenido de cada columna y se ajusta automáticamente el ancho de la columna, asegurando un ancho mínimo legible.

#### C. Modificación del Encabezado 'Valor' (Paso Crítico)
* Si se encontró la columna `'Valor'`, se modifica el texto del encabezado en la hoja de Excel a **"Valor (%)"**.

3.  **Guardado Final:** El libro de trabajo de `openpyxl` se guarda con todos los estilos aplicados, sobrescribiendo el archivo temporal en `OUTPUT_PATH`.

## 📌 Requisitos

Para que este script funcione, necesitas tener instaladas las siguientes librerías de Python:

* `pandas`
* `openpyxl`
* `numpy` (aunque se importa, no se utiliza directamente en el flujo principal visible).

```bash
pip install pandas openpyxl numpy



Pseudocodigo:

// --- CONSTANTES GLOBALES (Estilos) ---
CONST AZUL_RELLENO = "Color Azul"
CONST FUENTE_NEGRITA = "Negrita"
CONST ALINEACION_CENTRO = "Centrado"
CONST BORDE_FINO = "Borde Fino"
CONST ETIQUETA_TOTAL = "TOTAL"
CONST COLUMNA_VALOR_ESTANDAR = "Valor"

// =================================================================
// FUNCIÓN: limpiar_eliminar_total(ruta_entrada, ruta_salida, etiqueta_total)
// Propósito: Lee un Excel, elimina filas con la etiqueta 'TOTAL' en 'Cultivo' y guarda.
// =================================================================
FUNCION limpiar_eliminar_total(ruta_entrada, ruta_salida, etiqueta_total)
    LEER archivo Excel en ruta_entrada en DataFrame (DF)
    IMPRIMIR "Datos leídos de: " + ruta_entrada
    
    // Filtrado (Clave)
    DF_LIMPIO = FILTRAR DF DONDE Columna 'Cultivo' EN MAYÚSCULAS NO ES IGUAL A etiqueta_total
    
    GUARDAR DF_LIMPIO a Excel en ruta_salida (sin índice)
    IMPRIMIR "Cultivos '" + etiqueta_total + "' eliminados."
    
    RETORNAR DF_LIMPIO
FIN FUNCION

// =================================================================
// FUNCIÓN: aplicar_estetica_y_guardar(DF_limpio, ruta_salida)
// Propósito: Carga el archivo recién guardado y le aplica formato Excel avanzado.
// =================================================================
FUNCION aplicar_estetica_y_guardar(DF_limpio, ruta_salida)
    INTENTAR
        CARGAR libro de trabajo (WB) desde ruta_salida
        WS = Hoja Activa
        
        // 1. Encontrar columna 'Valor'
        valor_col_index = BUSCAR índice de columna 'Valor' en DF_limpio (+1 para Excel)
        SI valor_col_index es -1 ENTONCES
            IMPRIMIR "Advertencia: Columna 'Valor' no encontrada."
        FIN SI

        // 2. Aplicar Estilos a Encabezados (Fila 1)
        PARA CADA columna en Fila 1 de WS HACER
            PARA CADA celda en columna HACER
                APLICAR estilos: FUENTE_NEGRITA, AZUL_RELLENO, ALINEACION_CENTRO, BORDE_FINO a celda
            FIN PARA
        FIN PARA
        
        // 3. Ajustar ancho y aplicar estilos a datos
        PARA CADA columna en WS HACER
            max_length = 0
            column_letter = Letra de la columna
            
            PARA CADA celda_index, celda en columna HACER
                max_length = MAX(max_length, LONGITUD(Valor de la celda))
                celda.BORDE = BORDE_FINO
                SI celda_index > 0 ENTONCES // No es encabezado
                    celda.ALINEACION = ALINEACION_CENTRO
                FIN SI
            FIN PARA
            
            // Establecer Ancho
            ANCHO = MAX(max_length + 2, 12)
            ESTABLECER ancho de column_letter a ANCHO
        FIN PARA

        // 4. Modificar encabezado 'Valor' (Paso Crítico)
        SI valor_col_index NO ES -1 ENTONCES
            Celda(Fila 1, valor_col_index).Valor = "Valor (%)"
        FIN SI

        // 5. Guardar
        GUARDAR WB en ruta_salida
        IMPRIMIR "✅ ¡Éxito! Archivo limpiado y formateado."

    EXCEPTO Excepcion E:
        IMPRIMIR "❌ Ocurrió un error al aplicar la estética: " + E
    FIN INTENTAR
FIN FUNCION

// =================================================================
// FUNCIÓN: generar_grafico(DF_data, nombre_cultivo, ruta_salida)
// Propósito: Genera y guarda un gráfico de líneas (Valor vs Año).
// =================================================================
FUNCION generar_grafico(DF_data, nombre_cultivo, ruta_salida)
    DF_data = ESTABLECER Columna 'Año' como Índice
    
    CREAR Figura (tamaño 10x6)
    
    // Crear el gráfico de líneas
    GRAFICAR línea de DF_data.ÍNDICE (Año) vs Columna 'Valor (%)'
    
    ESTABLECER Título, Etiquetas (Eje X: Año, Eje Y: Valor (%))
    MOSTRAR cuadrícula
    AJUSTAR Eje X (Años) con rotación 45 grados

    GUARDAR Figura en ruta_salida
    MOSTRAR Gráfico
    CERRAR Figura
    
    IMPRIMIR "✅ Gráfico guardado en: " + ruta_salida
FIN FUNCION

// =================================================================
// FUNCIÓN: filtrar_y_exportar_cultivo(ruta_entrada, dir_salida)
// Propósito: Pide cultivo, filtra, exporta y grafica.
// =================================================================
FUNCION filtrar_y_exportar_cultivo(ruta_entrada, dir_salida)
    INTENTAR
        // 1. Lectura y Normalización
        DF = LEER Excel desde ruta_entrada
        IMPRIMIR "Base de datos leída."
        
        RENOMBRAR columna que contenga 'VALOR' y '%' a COLUMNA_VALOR_ESTANDAR ('Valor')
        
        SOLICITAR a USUARIO: nombre_cultivo (exacto)
        
        // 2. Filtrado y Ordenamiento
        cultivo_limpio = nombre_cultivo.QUITAR_ESPACIOS
        DF_FILTRADO = FILTRAR DF DONDE Columna 'Cultivo' es igual a cultivo_limpio
        
        SI DF_FILTRADO ESTÁ VACÍO ENTONCES
            IMPRIMIR "❌ No se encontraron datos para el cultivo."
            RETORNAR
        FIN SI
        
        ORDENAR DF_FILTRADO por 'Año'

        // 3. Generación de Archivo y Gráfico
        RENOMBRAR columna COLUMNA_VALOR_ESTANDAR a 'Valor (%)' (para salida)
        
        ruta_excel = GENERAR ruta y nombre para archivo Cultivo_<cultivo>.xlsx
        ruta_grafico = GENERAR ruta y nombre para archivo Grafico_<cultivo>.png

        GUARDAR DF_FILTRADO a Excel en ruta_excel (sin índice)
        IMPRIMIR "Datos filtrados guardados inicialmente."
        
        LLAMAR generar_grafico(DF_FILTRADO, nombre_cultivo, ruta_grafico)
        
        // 4. Aplicación de Estilos a Excel Filtrado
        WB = CARGAR libro de trabajo desde ruta_excel
        WS = Hoja Activa
        valor_col_index = BUSCAR índice de columna 'Valor (%)' en DF_FILTRADO (+1)
        
        // (Aplicar los mismos estilos de encabezado, bordes, centrado y ancho de columna que en aplicar_estetica_y_guardar)

        GUARDAR WB en ruta_excel
        IMPRIMIR "✅ ¡Proceso Completo! Archivo Excel y Gráfico guardados."

    EXCEPTO Excepcion E:
        IMPRIMIR "❌ Error: " + E
    FIN INTENTAR
FIN FUNCION

// =================================================================
// PROGRAMA PRINCIPAL (_main_)
// =================================================================
PROGRAMA PRINCIPAL
    // --- FASE 1: Limpieza General (Usando rutas de ejemplo) ---
    ruta_entrada_f1 = '/content/drive/.../Partefinal1.xlsx'
    ruta_salida_f1 = '/content/drive/.../Finalreto1.xlsx'
    
    INTENTAR
        DF_LIMPIO = LLAMAR limpiar_eliminar_total(ruta_entrada_f1, ruta_salida_f1, ETIQUETA_TOTAL)
        
        SI DF_LIMPIO NO ES NULO ENTONCES
            LLAMAR aplicar_estetica_y_guardar(DF_LIMPIO, ruta_salida_f1)
        FIN SI
    EXCEPTO Excepcion E:
        IMPRIMIR "❌ El proceso de limpieza falló: " + E
    FIN INTENTAR

    // --- FASE 2: Filtrado Específico y Graficado (Usando la salida de FASE 1 como entrada) ---
    ruta_entrada_f2 = ruta_salida_f1 // '/content/drive/.../Finalreto1.xlsx'
    dir_salida_f2 = '/content/drive/.../Final'
    
    LLAMAR filtrar_y_exportar_cultivo(ruta_entrada_f2, dir_salida_f2)

FIN PROGRAMA
