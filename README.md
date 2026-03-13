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
