# Integración Excel-PostgreSQL con xlwings

## 📋 Objetivo del Proyecto

Este es un proyecto de **prueba de concepto** que demuestra la integración entre Microsoft Excel y una base de datos PostgreSQL utilizando xlwings. El objetivo es permitir que los usuarios de Excel puedan ejecutar funciones personalizadas que consulten datos directamente desde la base de datos PostgreSQL.

### Funcionalidad Principal
- **Función personalizada en Excel**: `get_stock_data()` que consulta el stock de productos desde PostgreSQL
- **Integración transparente**: Los usuarios pueden usar la función directamente en las celdas de Excel
- **Manejo de errores**: La función devuelve mensajes de error comprensibles en caso de problemas

## ⚠️ Problema Actual - Error de Acceso Denegado

### Descripción del Error
Al hacer clic en el botón **"Import Functions"** en la cinta de xlwings de Excel, se recibe el mensaje:
```
Acceso denegado
```

### Causa del Problema
Este error se debe a:
1. **Políticas corporativas** que restringen la ejecución de scripts
2. **Falta de permisos de administrador** para registrar el complemento de xlwings
3. **Configuración de seguridad de Windows** que bloquea la ejecución de archivos .dll

### Soluciones Recomendadas
1. **Ejecutar Excel como Administrador** (solución más común)
2. **Configurar políticas de grupo** para permitir xlwings
3. **Registrar manualmente el complemento** usando el registro de Windows

## 🛠️ Prerrequisitos

Antes de comenzar, asegúrate de tener instalado:

### Software Requerido
- **Python 3.8 o superior** ([Descargar Python](https://www.python.org/downloads/))
- **Microsoft Excel** (2016 o superior)
- **Git** ([Descargar Git](https://git-scm.com/downloads))

### Verificar Instalaciones
Abre la terminal (cmd o PowerShell) y ejecuta:
```bash
python --version
pip --version
git --version
```

Si alguno de estos comandos no funciona, instala el software correspondiente.

## 📥 Instalación del Proyecto

### Paso 1: Clonar el Repositorio
```bash
# Navegar al directorio donde quieres instalar el proyecto
cd C:\Users\[tu_usuario]\Documents

# Clonar el repositorio
git clone git@github.com:rjardi/py-xls-integration.git
cd py-xls-integration
```

### Paso 2: Crear Entorno Virtual
```bash
# Crear entorno virtual
python -m venv venv

# Activar entorno virtual (Windows)
venv\Scripts\activate

# Verificar que el entorno está activo (debería mostrar (venv) al inicio)
```

### Paso 3: Instalar Dependencias
```bash
# Asegúrate de que el entorno virtual esté activo
pip install -r requirements.txt
```

### Paso 4: Configurar Variables de Entorno
1. **Crear archivo .env**:
   ```bash
   # Copiar el archivo de ejemplo
   copy .env.example .env
   ```

2. **Editar el archivo .env** con tus datos de conexión:
   ```env
   DB_HOST_URL=tu_servidor_postgresql
   DB_NAME=nombre_base_datos
   DB_USER=tu_usuario
   DB_PASSWORD=tu_contraseña
   ```

### Paso 5: Probar la Conexión
```bash
# Ejecutar script de prueba
python test_connection.py
```

Si todo está correcto, deberías ver:
```
✓ Conexión exitosa a PostgreSQL
✓ Función ejecutada correctamente
```

## 🔧 Configuración de xlwings

### Paso 1: Registrar el Complemento
1. **Cerrar Excel completamente**
2. **Abrir terminal como Administrador**:
   - Presiona `Windows + X`
   - Selecciona "Windows PowerShell (Administrador)" o "Símbolo del sistema (Administrador)"

3. **Navegar al proyecto y activar entorno virtual**:
   ```bash
   cd C:\ruta\a\tu\proyecto\py-xls-integration
   venv\Scripts\activate
   ```

4. **Registrar xlwings**:
   ```bash
   xlwings addin install
   ```

### Paso 2: Configurar Excel
1. **Abrir Excel**
2. **Ir a la pestaña "xlwings"** en la cinta
3. **Hacer clic en "Import Functions"**
4. **Si aparece error de acceso denegado**:
   - Cerrar Excel
   - Abrir Excel como Administrador
   - Repetir el proceso

### Paso 3: Probar la Función
1. **Abrir el archivo** `excel_test.xlsm`
2. **En una celda**, escribir:
   ```
   =get_stock_data("PDX", "BRAM", "22-09-2025")
   ```
3. **Presionar Enter**

## 🚨 Solución de Problemas

### Error: "Acceso denegado" al importar funciones

#### Solución 1: Ejecutar como Administrador
1. Cerrar Excel
2. Hacer clic derecho en Excel
3. Seleccionar "Ejecutar como administrador"
4. Repetir el proceso de importación

#### Solución 2: Registrar manualmente
```bash
# En terminal como administrador
cd C:\ruta\a\tu\proyecto\py-xls-integration
venv\Scripts\activate
xlwings addin install --force
```

#### Solución 3: Verificar políticas de grupo
1. Presionar `Windows + R`
2. Escribir `gpedit.msc`
3. Navegar a: `Configuración del equipo > Plantillas administrativas > Sistema`
4. Buscar "Ejecutar scripts de Windows PowerShell"
5. Configurar como "Habilitado" o "No configurado"

### Error: "No se puede conectar a la base de datos"
1. Verificar que el archivo `.env` existe y tiene los datos correctos
2. Probar la conexión con `python test_connection.py`
3. Verificar que el servidor PostgreSQL esté accesible

### Error: "Módulo no encontrado"
1. Verificar que el entorno virtual esté activo
2. Reinstalar dependencias: `pip install -r requirements.txt`

## 📁 Estructura del Proyecto

```
py-xls-integration/
├── README.md                 # Este archivo
├── requirements.txt          # Dependencias de Python
├── .env.example             # Plantilla de variables de entorno
├── .env                     # Variables de entorno (crear manualmente)
├── get_stock_data.py        # Función principal de xlwings
├── test_connection.py       # Script de prueba de conexión
├── excel_test.xlsm          # Archivo de Excel de prueba
└── venv/                    # Entorno virtual de Python
```

## 🔍 Archivos del Proyecto

### `get_stock_data.py`
- Contiene la función `get_stock_data()` que se ejecuta en Excel
- Se conecta a PostgreSQL y ejecuta la función `api_xls.f_pla_qty_stock()`
- Maneja errores y devuelve resultados a Excel

### `test_connection.py`
- Script de prueba para verificar la conexión a la base de datos
- Prueba la función de stock con parámetros de ejemplo
- Útil para diagnosticar problemas antes de usar Excel

### `excel_test.xlsm`
- Archivo de Excel con macros habilitadas
- Contiene ejemplos de uso de la función `get_stock_data()`

## 📞 Soporte

Si encuentras problemas:

1. **Verificar logs**: Revisar la salida de `test_connection.py`
2. **Comprobar permisos**: Asegurar que Excel se ejecuta con permisos adecuados
3. **Revisar configuración**: Verificar que el archivo `.env` esté correctamente configurado
4. **Contactar administrador**: Para problemas de políticas corporativas

## 📝 Notas Importantes

- **Seguridad**: Nunca compartas el archivo `.env` ya que contiene credenciales
- **Backup**: Mantén una copia de seguridad de tu configuración
- **Actualizaciones**: Actualiza las dependencias regularmente con `pip install -r requirements.txt --upgrade`
- **Logs**: Los errores se muestran directamente en las celdas de Excel

---

**Versión**: 1.0  
**Última actualización**: Enero 2025  
**Compatibilidad**: Excel 2016+, Python 3.8+
