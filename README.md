# 📝 Generador de Actas de Operatividad

Una aplicación de escritorio con interfaz gráfica para generar documentos de actas de operatividad de manera automatizada, con capacidad de exportación a PDF.

## ✨ Características

- 🖥️ **Interfaz gráfica moderna** con tkinter
- 📄 **Soporte para plantillas 3G y 4G**
- 📝 **Reemplazo automático de placeholders**
- 📁 **Selector de carpeta de destino**
- 🔄 **Exportación a PDF** (requiere Microsoft Word)
- 🎨 **Preservación de formato** (incluyendo WordArt)
- 📊 **Autocompletado de fechas**
- 🚀 **Compilación a ejecutable**

## 🛠️ Instalación

### Requisitos del Sistema

- Python 3.7+
- Microsoft Word (para exportación a PDF)
- Windows 10/11 (recomendado)

### Instalación Rápida

```bash
# Clonar el repositorio
git clone https://github.com/tuusuario/generador-actas.git
cd generador-actas

# Instalar dependencias básicas
pip install python-docx pywin32

# O instalar todas las dependencias
pip install -r requirements.txt
```

## 🚀 Uso

### Ejecutar la Aplicación

```bash
python fill_templates.py
```

### Pasos para Generar un Acta

1. **Seleccionar plantilla**: Elige entre 3G o 4G
2. **Completar campos**: Llena los datos requeridos
3. **Elegir destino**: Selecciona dónde guardar el archivo
4. **Generar**: Crea el documento DOCX y/o PDF

### Campos Disponibles

- **NUMERO**: Número del acta
- **EMPRESA**: Nombre de la empresa
- **RUC**: RUC de la empresa
- **PLACA**: Placa del vehículo
- **IMEI**: IMEI del dispositivo
- **FECHA DE INSTALACION**: Fecha de instalación
- **DIA**: Día (autocompletado)
- **MES**: Mes (autocompletado)

## 📁 Estructura del Proyecto

```
generador-actas/
├── fill_templates.py          # Aplicación principal
├── requirements.txt           # Dependencias
├── README.md                 # Documentación
├── .gitignore               # Archivos excluidos
├── TEMPLATE_GUIDE.md        # Guía de plantillas
└── export/                  # Scripts de compilación
    ├── build_simple.bat     # Compilación básica
    ├── build_pyinstaller.bat # Compilación con PyInstaller
    └── GeneradorActas.spec  # Configuración PyInstaller
```

## 🔧 Configuración de Plantillas

### Formato de Placeholders

Las plantillas deben contener placeholders con el formato `{{nombre}}`:

```
{{numero}} - {{empresa}}
RUC: {{ruc}}
Placa: {{placa}}
IMEI: {{imei}}
Fecha: {{fec_ins}}
Día {{dia}} de {{mes}}
```

### Ubicación de Plantillas

- Crea tus plantillas en formato `.docx`
- Colócalas en la misma carpeta que `fill_templates.py`
- Nómbralas como `MOLDE 3G.docx` y `MOLDE 4G.docx`

> **⚠️ Importante**: Las plantillas no están incluidas en el repositorio por motivos de seguridad. Consulta `TEMPLATE_GUIDE.md` para crear las tuyas.

## 📦 Crear Ejecutable

### Usando Nuitka (Recomendado)

```bash
# Instalar Nuitka
pip install nuitka

# Ejecutar script de compilación
./export/build_simple.bat
```

### Usando PyInstaller

```bash
# Instalar PyInstaller
pip install pyinstaller

# Compilar
./export/build_pyinstaller.bat
```

El ejecutable se generará en la carpeta `dist/`.

## 🔍 Solución de Problemas

### Error: "No se encontró la plantilla"

- Verifica que las plantillas estén en la carpeta correcta
- Asegúrate de que tengan los nombres exactos: `MOLDE 3G.docx` y `MOLDE 4G.docx`

### Error: "PDF export no disponible"

- Instala pywin32: `pip install pywin32`
- Verifica que Microsoft Word esté instalado
- Ejecuta el script como administrador si es necesario

### Error: "No se detectaron reemplazos"

- Verifica que los placeholders en la plantilla sean exactos: `{{nombre}}`
- Asegúrate de que no haya espacios extra en los placeholders

## 🤝 Contribuir

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/nueva-funcionalidad`)
3. Commit tus cambios (`git commit -am 'Agrega nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/nueva-funcionalidad`)
5. Crea un Pull Request

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Ver el archivo `LICENSE` para más detalles.

## 🎯 Características Futuras

- [ ] Soporte para más tipos de plantillas
- [ ] Exportación a otros formatos (ODT, RTF)
- [ ] Interfaz web opcional
- [ ] Procesamiento por lotes
- [ ] Plantillas personalizables desde la interfaz

---

⭐ Si este proyecto te ha sido útil, ¡considera darle una estrella!
