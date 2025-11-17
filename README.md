# Integración Excel-PostgreSQL con VBA + ODBC

## 📋 Índice
1. [Objetivo del Proyecto](#objetivo-del-proyecto)
2. [Prerrequisitos](#prerrequisitos)
3. [Instalación Paso a Paso](#instalación-paso-a-paso)
4. [Configuración de Excel](#configuración-de-excel)
5. [Prueba de Funcionamiento](#prueba-de-funcionamiento)
6. [Solución de Problemas](#solución-de-problemas)
7. [Estructura del Proyecto](#estructura-del-proyecto)
8. [Notas Adicionales](#notas-adicionales)

---

## 🎯 Objetivo del Proyecto

Este proyecto permite **conectar Excel directamente con PostgreSQL** usando **VBA + ODBC** sin necesidad de Python. La solución incluye:

- **Funciones personalizadas** para consultar datos de stock, entradas y salidas desde PostgreSQL
- **Conexión persistente**: Reutiliza la conexión para mejorar rendimiento
- **Sin permisos de administrador**: Funciona con políticas corporativas restrictivas
- **Fácil instalación**: Pasos simples para cualquier usuario

---

## 🛠️ Prerrequisitos

### Software Requerido
- **Microsoft Excel** (2016 o superior)
- **Git** ([Descargar Git](https://git-scm.com/downloads))
- **Permisos de administrador** (solo para instalar driver ODBC)

### Verificar Instalaciones
Abre la terminal (cmd o PowerShell) y ejecuta:
```bash
git --version
```

---

## 📥 Instalación Paso a Paso

### Paso 1: Clonar el Repositorio
```bash
# Navegar al directorio deseado
cd C:\Users\[tu_usuario]\Documents

# Clonar el repositorio
git clone https://github.com/rjardi/py-xls-integration.git
cd py-xls-integration
```

### Paso 2: Instalar Driver ODBC PostgreSQL

#### 2.1 Verificar Arquitectura de Excel
1. **Abrir Excel**
2. **Ir a Archivo > Cuenta > Acerca de Excel**
3. **Verificar si es "32-bit" o "64-bit"**

#### 2.2 Descargar Driver Correcto
1. **Ir a**: [PostgreSQL ODBC Releases](https://www.postgresql.org/ftp/odbc/releases/)
2. **Descargar la última versión** (ej: `REL-17_00_0006`)
3. **Seleccionar archivo .msi según arquitectura**:
   - **64-bit**: `psqlodbc-17.00.0006-x64.msi`
   - **32-bit**: `psqlodbc-17.00.0006-x86.msi`

#### 2.3 Instalar Driver
1. **Ejecutar el archivo .msi como Administrador**
2. **Seguir el asistente de instalación**
3. **Verificar instalación**:
   - Presionar `Windows + R`
   - Escribir `odbcad32.exe`
   - Verificar que aparece "PostgreSQL Unicode" en la lista

### Paso 3: Configurar DSN de Usuario en Windows

#### 3.1 Abrir Administrador de Origen de Datos ODBC
1. **Presionar `Windows + R`**
2. **Escribir**: `odbcad32.exe`
3. **Presionar Enter**
4. **Seleccionar pestaña "DSN de usuario"** (User DSN)

#### 3.2 Crear Nuevo DSN
1. **Clic en "Agregar"**
2. **Seleccionar "PostgreSQL Unicode"** de la lista
3. **Clic en "Finalizar"**

#### 3.3 Configurar Parámetros de Conexión
En la ventana de configuración, completar los siguientes campos:

| Campo | Valor | Descripción |
|-------|-------|-------------|
| **Data Source** | `PostgreSQL35W` | Nombre del DSN (puedes cambiarlo) |
| **Server** | `tu_servidor_postgresql` | IP o nombre del servidor PostgreSQL |
| **Port** | `5432` | Puerto de PostgreSQL (por defecto) |
| **Database** | `tu_base_datos` | Nombre de la base de datos |
| **Username** | `tu_usuario` | Usuario de PostgreSQL |
| **Password** | `tu_contraseña` | Contraseña del usuario |
| **SSLMode** | `require` | Modo SSL requerido

#### 3.5 Guardar Configuración
1. **Clic en "OK"** para guardar
2. **Verificar** que aparece `PostgreSQL35W` en la lista de DSN de usuario
3. **Cerrar** el Administrador ODBC

> **Nota**: No necesitas permisos de administrador para crear DSN de usuario, solo para instalar el driver ODBC.

---

## 🔧 Configuración de Excel

### Paso 1: Importar Código VBA
1. **Abrir Excel**
2. **Presionar `ALT + F11`** (abrir editor VBA)
3. **En el menú**: `Archivo > Importar archivo`
4. **Seleccionar**: `get_stock_data.vba` del proyecto
5. **Repetir el paso 4** para `get_stock_data_sem.vba` (funciones semanales de valor único)
6. **Repetir el paso 4** para `get_stock_data_grj.vba` (función por granja/lote/nave/art)
7. **Repetir el paso 4** para `set_stock_data.vba` (funciones de tabla 2D)
8. **Cerrar el editor VBA**

### Paso 2: Habilitar Macros
1. **Guardar el archivo como `.xlsm`** (Excel con macros)
2. **Si aparece advertencia de seguridad**: Hacer clic en "Habilitar contenido"

---

## ✅ Prueba de Funcionamiento

### Prueba Básica (Stock por fecha exacta)
1. **En cualquier celda de Excel**, escribir:
   ```
   =GetStockAvi("PDX";"STOCK_QTY";"BRAM";"2025-10-10";0;999;0;99,999)
   ```
2. **Resultado**: Cantidad (u otro valor) devuelto por la función SQL `api_xls.f_pla_get_data_stock_v1`

### Prueba de Entradas (rango de fechas)
1. **En una celda**, escribir:
   ```
   =GetEntradaAvi("PDX";"ENTRADAS_QTY";"BRAM";"2025-10-01";"2025-10-09")
   ```
2. **Resultado**: Valor calculado para entradas en el rango

### Prueba de Salidas (rango + filtros)
1. **En una celda**, escribir:
   ```
   =GetSalidasAvi("PDX";"SALIDAS_QTY";"BRAM";"2025-10-01";"2025-10-09";0;999;0;99,999)
   ```
2. **Resultado**: Valor calculado para salidas con los filtros indicados

### Pruebas Semanales (valor único)
- **Stock semanal**
  ```
  =GetStockAviSem("PDX";"STOCK_QTY";"BRAM";2025;44;0;999;0;9999)
  ```
- **Entradas semanales**
  ```
  =GetEntradaAviSem("PDX";"ENTRADAS_QTY";"BRAM";2025;44)
  ```
- **Salidas semanales**
  ```
  =GetSalidasAviSem("PDX";"SALIDAS_QTY";"BRAM";2025;44;0;999;0;9999)
  ```

### Prueba por Granja/Lote/Nave/Artículo
1. **En una celda**, escribir:
   ```
   =GetStockAviGrj("PDX";"STOCK_QTY";"2025-10-10";"20378T";"39";"4";"3001")
   ```
2. **Resultado**: Valor numérico de la petición para la granja y artículo indicados

### Tip: Forzar recálculo de fórmulas en Excel
Si las fórmulas ya fueron calculadas y quieres actualizar los resultados, presiona:
```
Ctrl + Alt + F9
```
Esto fuerza el recálculo completo del libro.

### Prueba de Conexión
1. **En una celda**, escribir:
   ```
   =TestConnection()
   ```
2. **Resultado esperado**: "Conexión exitosa"

### Prueba SSL
1. **En una celda**, escribir:
   ```
   =TestConnectionSSL()
   ```
2. **Resultado esperado**: "Conexión SSL exitosa"

---

## 📊 Funciones de Tabla (Array 2D)

Las funciones `GetSet*` devuelven **tablas completas** en lugar de valores únicos. Excel muestra automáticamente los datos en formato de tabla con encabezados y filas de datos.

### Prueba de Tabla de Stock
1. **En una celda**, escribir:
   ```
   =GetSetStockAvi("PDX";"GRJ_STOCK_QTY";"Broiler Amarillo";"2025-01-15";0;99;0;99,999)
   ```
2. **Resultado**: Tabla completa con encabezados (fecha_stock, articulo, peso_real, etc.) y filas de datos

### Prueba de Tabla de Entradas
1. **En una celda**, escribir:
   ```
   =GetSetEntradaAvi("PDX";"GRJ_ENTRADAS_QTY";"Broiler Amarillo";"2025-01-01";"2025-01-15")
   ```
2. **Resultado**: Tabla con datos de entradas en el rango de fechas

### Prueba de Tabla de Salidas
1. **En una celda**, escribir:
   ```
   =GetSetSalidasAvi("PDX";"GRJ_SALIDAS_QTY";"Broiler Amarillo";"2025-01-01";"2025-01-15";0;99;0;99,999)
   ```
2. **Resultado**: Tabla con datos de salidas con filtros aplicados


### 💡 Convertir a Tabla de Excel
Para aplicar formato de tabla profesional:
1. **Seleccionar el rango** que contiene los datos
2. **Presionar `Ctrl + T`** (Insertar > Tabla)
3. **Marcar "Mi tabla tiene encabezados"**
4. **Elegir estilo** y aceptar

### ⚡ Ventajas de las Funciones de Tabla
- ✅ **Sin restricciones de UDF**: Funcionan desde cualquier celda
- ✅ **Formato automático**: Excel muestra la tabla inmediatamente
- ✅ **Datos completos**: Acceso a todos los campos del JSON
- ✅ **Fácil análisis**: Datos listos para gráficos y análisis

---

## 🚨 Solución de Problemas

### Error: "No se encuentra el nombre del origen de datos"
**Causa**: DSN no configurado correctamente
**Solución**:
1. Verificar que el DSN `PostgreSQL35W` existe en "DSN de usuario"
2. Abrir `odbcad32.exe` y verificar configuración
3. Reinstalar driver con arquitectura correcta si es necesario
4. Reiniciar Excel

### Error: "could not translate host name"
**Causa**: Servidor no accesible
**Solución**:
1. Verificar conexión a internet/VPN
2. Comprobar IP del servidor en la configuración del DSN
3. Probar con IP en lugar de nombre de servidor
4. Usar "Test" en la configuración del DSN

### Error: "failed no pg_hba.conf entry"
**Causa**: Problema de autenticación PostgreSQL
**Solución**:
1. Verificar usuario y contraseña en la configuración del DSN
2. Asegurar que "SSL Mode" está configurado como "Require"
3. Contactar administrador de base de datos

### Error: "Object variable not set"
**Causa**: Problema en código VBA o DSN
**Solución**:
1. Verificar que el DSN `PostgreSQL35W` está configurado correctamente
2. Probar con `TestConnection()` para verificar conectividad
3. Revisar que el nombre del DSN coincide exactamente en el código VBA

---

## 📁 Estructura del Proyecto

```
py-xls-integration/
├── README.md                    # Este archivo
├── get_stock_data.vba          # Funciones de valor único (fecha/rango)
├── get_stock_data_sem.vba      # Funciones de valor único (semanas)
├── get_stock_data_grj.vba      # Función de valor único por granja/lote/nave/art
├── set_stock_data.vba          # Funciones de tabla (Array 2D)
├── test_functions.vba          # Funciones de prueba
├── postgresql.dsn.example      # Plantilla de configuración (referencia)
├── odbc_driver/                # Driver ODBC portable (opcional)
│   ├── psqlodbc35w.dll
│   └── libpq.dll
└── python/                     # Solución Python (paralizada)
    ├── get_stock_data.py
    ├── test_connection.py
    └── requirements.txt
```

---

## 📝 Notas Adicionales

### Funciones Disponibles

#### 🔢 Funciones de Valor Único
- `GetStockAvi(unidad_operacional, peticion, producto_venta, fecha_dato, DiaVida_inicial, DiaVida_final, peso_inicial, peso_final)`
  - **PostgreSQL**: `api_xls.f_pla_get_data_stock_v1`
  - **Devuelve**: Decimal (valor único)
- `GetEntradaAvi(unidad_operacional, peticion, producto_venta, fch_inicial, fch_final)`
  - **PostgreSQL**: `api_xls.f_pla_get_data_stock_v1`
  - **Devuelve**: Decimal (valor único)
- `GetSalidasAvi(unidad_operacional, peticion, producto_venta, fch_inicial, fch_final, DiaVida_inicial, DiaVida_final, peso_inicial, peso_final)`
  - **PostgreSQL**: `api_xls.f_pla_get_data_stock_v1`
  - **Devuelve**: Decimal (valor único)
- `GetStockAviSem(unidad_operacional, peticion, producto_venta, año, semana, DiaVida_inicial, DiaVida_final, peso_inicial, peso_final)`
  - **PostgreSQL**: `api_xls.f_pla_get_data_stock_sem_v1`
  - **Devuelve**: Decimal (valor único semanal)
- `GetEntradaAviSem(unidad_operacional, peticion, producto_venta, año, semana)`
  - **PostgreSQL**: `api_xls.f_pla_get_data_stock_sem_v1`
  - **Devuelve**: Decimal (valor único semanal)
- `GetSalidasAviSem(unidad_operacional, peticion, producto_venta, año, semana, DiaVida_inicial, DiaVida_final, peso_inicial, peso_final)`
  - **PostgreSQL**: `api_xls.f_pla_get_data_stock_sem_v1`
  - **Devuelve**: Decimal (valor único semanal)
- `GetStockAviGrj(unidad_operacional, peticion, fecha_dato, granja, lote, nave, articulo)`
  - **PostgreSQL**: `api_xls.f_pla_get_data_stock_grj_v1`
  - **Devuelve**: Decimal (valor único por granja/lote/nave/artículo)

#### 📊 Funciones de Tabla (Array 2D)
- `GetSetStockAvi(unidad_operacional, peticion, producto_venta, fecha_dato, DiaVida_inicial, DiaVida_final, peso_inicial, peso_final)`
  - **PostgreSQL**: `api_xls.f_pla_get_set_data_stock_v1`
  - **Devuelve**: Array 2D (tabla con múltiples filas y columnas)
- `GetSetEntradaAvi(unidad_operacional, peticion, producto_venta, fch_inicial, fch_final)`
  - **PostgreSQL**: `api_xls.f_pla_get_set_data_stock_v1`
  - **Devuelve**: Array 2D (tabla con múltiples filas y columnas)
- `GetSetSalidasAvi(unidad_operacional, peticion, producto_venta, fch_inicial, fch_final, DiaVida_inicial, DiaVida_final, peso_inicial, peso_final)`
  - **PostgreSQL**: `api_xls.f_pla_get_set_data_stock_v1`
  - **Devuelve**: Array 2D (tabla con múltiples filas y columnas)

#### 🔧 Funciones de Utilidad
- `TestConnection()`: Prueba conexión básica
- `TestConnectionSSL()`: Prueba conexión con SSL
- `InitializeConnection()`: Inicializa conexión persistente
- `CloseGlobalConnection()`: Cierra conexión global

### Funciones de PostgreSQL Utilizadas
Todas las funciones de PostgreSQL utilizan el sufijo `_v1`:
- `api_xls.f_pla_get_data_stock_v1`: Funciones diarias (valor único)
- `api_xls.f_pla_get_data_stock_sem_v1`: Funciones semanales (valor único, sin parámetro `p_fecha_dato`)
- `api_xls.f_pla_get_data_stock_grj_v1`: Función por granja/lote/nave/artículo (valor único)
- `api_xls.f_pla_get_set_data_stock_v1`: Funciones de tabla (devuelve JSON parseado como Array 2D)

### Optimizaciones Incluidas
- **Conexión persistente**: Reutiliza la misma conexión para mejorar rendimiento
- **Manejo de errores**: Mensajes claros en caso de problemas
- **Debug**: Usa `Debug.Print` para ver logs en ventana inmediata
- **Formato de fechas**: Todas las fechas se envían como strings en formato `yyyy-mm-dd` (adVarChar)
- **Conversión numérica**: Los valores numéricos del JSON se convierten automáticamente de punto a coma decimal

### Seguridad
- **DSN de Usuario**: Credenciales almacenadas de forma segura en Windows
- **SSL**: Conexión encriptada configurada en el DSN
- **Sin hardcoding**: Credenciales gestionadas por el sistema operativo
- **Permisos de usuario**: No requiere permisos de administrador para funcionar

---

## 🔄 Alternativa Python (Paralizada)

Este repositorio también contiene una solución usando **Python + xlwings** en la carpeta `python/`, pero está **paralizada** debido a problemas de permisos corporativos. El error "Acceso denegado" al importar funciones desde el add-in de xlwings impide su uso en entornos con políticas de seguridad restrictivas.

**Ventajas de la solución VBA actual**:
- ✅ No requiere permisos de administrador para funcionar
- ✅ No depende de Python instalado
- ✅ Funciona con políticas corporativas restrictivas
- ✅ Instalación más simple para usuarios finales

---

**Versión**: 2.0  
**Última actualización**: Enero 2025  
**Compatibilidad**: Excel 2016+, Windows 10+