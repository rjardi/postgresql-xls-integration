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

- **Función personalizada**: `GetStockData()` que consulta stock desde PostgreSQL
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

### Paso 3: Crear Archivo de Configuración DSN

#### 3.1 Crear archivo `postgresql.dsn`
En la raíz del proyecto, crear archivo `postgresql.dsn` con este contenido:

```ini
[ODBC]
DRIVER=PostgreSQL Unicode
SERVER=tu_servidor_postgresql
DATABASE=tu_base_datos
UID=tu_usuario
PWD=tu_contraseña
PORT=5432
SSLmode=require
```

#### 3.2 Reemplazar Valores
- `tu_servidor_postgresql`: IP o nombre del servidor
- `tu_base_datos`: Nombre de la base de datos
- `tu_usuario`: Usuario de PostgreSQL
- `tu_contraseña`: Contraseña del usuario

---

## 🔧 Configuración de Excel

### Paso 1: Importar Código VBA
1. **Abrir Excel**
2. **Presionar `ALT + F11`** (abrir editor VBA)
3. **En el menú**: `Archivo > Importar archivo`
4. **Seleccionar**: `get_stock_data.vba` del proyecto
5. **Cerrar el editor VBA**

### Paso 2: Habilitar Macros
1. **Guardar el archivo como `.xlsm`** (Excel con macros)
2. **Si aparece advertencia de seguridad**: Hacer clic en "Habilitar contenido"

---

## ✅ Prueba de Funcionamiento

### Prueba Básica
1. **En cualquier celda de Excel**, escribir:
   ```
   =GetStockData("PDX", "BRAM", "22-09-2025")
   ```
2. **Presionar Enter**
3. **Resultado esperado**: Número de stock o mensaje de error

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

## 🚨 Solución de Problemas

### Error: "No se encuentra el nombre del origen de datos"
**Causa**: Driver ODBC no instalado correctamente
**Solución**:
1. Verificar que el driver aparece en `odbcad32.exe`
2. Reinstalar driver con arquitectura correcta
3. Reiniciar Excel

### Error: "could not translate host name"
**Causa**: Servidor no accesible
**Solución**:
1. Verificar conexión a internet/VPN
2. Comprobar IP del servidor en `postgresql.dsn`
3. Probar con IP en lugar de nombre

### Error: "failed no pg_hba.conf entry"
**Causa**: Problema de autenticación PostgreSQL
**Solución**:
1. Verificar usuario y contraseña en `postgresql.dsn`
2. Asegurar que `SSLmode=require` está presente
3. Contactar administrador de base de datos

### Error: "Object variable not set"
**Causa**: Problema en código VBA
**Solución**:
1. Verificar que el archivo `postgresql.dsn` existe
2. Comprobar que la ruta es correcta
3. Revisar permisos de lectura del archivo

---

## 📁 Estructura del Proyecto

```
py-xls-integration/
├── README.md                    # Este archivo
├── get_stock_data.vba          # Código VBA principal
├── postgresql.dsn              # Configuración ODBC (crear manualmente)
├── postgresql.dsn.example      # Plantilla de configuración
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
- `GetStockData(empavi, erpcodave, fecha)`: Función principal
- `TestConnection()`: Prueba conexión básica
- `TestConnectionSSL()`: Prueba conexión con SSL
- `InitializeConnection()`: Inicializa conexión persistente
- `CloseGlobalConnection()`: Cierra conexión global

### Optimizaciones Incluidas
- **Conexión persistente**: Reutiliza la misma conexión
- **Manejo de errores**: Mensajes claros en caso de problemas
- **Debug**: Usa `Debug.Print` para ver logs en ventana inmediata

### Seguridad
- **Archivo DSN**: Contiene credenciales, mantener privado
- **SSL**: Conexión encriptada con `SSLmode=require`
- **Sin hardcoding**: Credenciales en archivo externo

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