<p align="center"><a href="https://laravel.com" target="_blank"><img src="https://raw.githubusercontent.com/laravel/art/master/logo-lockup/5%20SVG/2%20CMYK/1%20Full%20Color/laravel-logolockup-cmyk-red.svg" width="400"></a></p>

<p align="center">
<a href="https://travis-ci.org/laravel/framework"><img src="https://travis-ci.org/laravel/framework.svg" alt="Build Status"></a>
<a href="https://packagist.org/packages/laravel/framework"><img src="https://img.shields.io/packagist/dt/laravel/framework" alt="Total Downloads"></a>
<a href="https://packagist.org/packages/laravel/framework"><img src="https://img.shields.io/packagist/v/laravel/framework" alt="Latest Stable Version"></a>
<a href="https://packagist.org/packages/laravel/framework"><img src="https://img.shields.io/packagist/l/laravel/framework" alt="License"></a>
</p>

# Automatizacion del formato de registro fotografico - AFR

Automatiza el registro de imágenes en documentos Word, agrupándolas y organizándolas según formato, con interfaz gráfica sencilla. Facilita y agiliza el trabajo que antes era realizado manualmente por mi abuela.

## 📋 Guía de Instalación y Uso

### **Requisitos Previos**

- Python 3.7 o superior instalado en tu sistema
- pip (gestor de paquetes de Python)

### **Instalación**

**1. Clonar el repositorio:**

```bash
git clone https://github.com/Keyver-Stiven/AFR.git
cd AFR
```

**2. Instalar las dependencias necesarias:**

```bash
pip install python-docx Pillow
```

### **Ejecución del Programa**

**Para ejecutar el programa directamente desde Python:**

```bash
python main.py
```

O si tienes Python 3 específicamente:

```bash
python3 main.py
```

### **Generación de Ejecutable con PyInstaller**

Si deseas crear un archivo ejecutable (.exe) para distribuir o usar sin necesidad de Python:

**1. Instalar PyInstaller:**

```bash
pip install pyinstaller
```

**2. Generar el ejecutable:**

```bash
pyinstaller --onefile --windowed main.py
```

**Opciones del comando:**
- `--onefile`: Genera un único archivo ejecutable
- `--windowed`: Ejecuta sin mostrar la consola (recomendado para aplicaciones con interfaz gráfica)

**3. Ubicación del ejecutable:**

El archivo .exe generado se encontrará en la carpeta `dist/` dentro del directorio del proyecto.

**Nota:** Si deseas personalizar el ícono del ejecutable, puedes agregar:

```bash
pyinstaller --onefile --windowed --icon=icono.ico main.py
```

---

## 🚀 Uso del Programa

Una vez iniciado el programa, sigue las instrucciones en la interfaz gráfica para seleccionar las imágenes y configurar el formato del documento Word que deseas generar.
