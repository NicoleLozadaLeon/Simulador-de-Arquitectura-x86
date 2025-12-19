# 🖥️ Simulador de Arquitectura x86
## 📖 Descripción del Proyecto 
Este proyecto consiste en el desarrollo de un simulador interactivo de arquitectura x86, diseñado para ilustrar el funcionamiento interno de los componentes clave de una computadora. El simulador permite cargar y ejecutar código en lenguaje ensamblador x86, visualizando en tiempo real el estado de la CPU, memoria, pipeline y otros elementos fundamentales.
## 🎯 Objetivo Principal
- Crear una herramienta educativa que facilite la comprensión de:
  
- La ejecución de instrucciones a nivel de hardware
  
- La gestión de memoria (RAM, Caché, Virtual)

- El funcionamiento del pipeline de instrucciones

- Las políticas de reemplazo en memoria caché

- La arquitectura Von Neumann vs Harvard

## 🛠️ Tecnologías Utilizadas
- Plataforma: Microsoft Excel con VBA (Visual Basic for Applications)

- Control de Versiones: GitHub con Git

- Gestión de Proyecto: GitHub Projects e Issues

- Documentación: GoogleDocs

## 📁 Estructura del Proyecto

Simulador-de-Arquitectura-x86/

│

├── 📂 CPU/                 # Simulación de Unidad de Control, ALU y Registros

├── 📂 grafico/             # Componentes de visualización e interfaz

├── 📂 memoria/             # Gestión de RAM, Caché y Memoria Virtual

├── 📂 pipeline/            # Simulación del pipeline de instrucciones

├── 📂 traductor/           # Traducción de código (funcionalidad opcional)

├── 📄 programa.cls         # Clase principal del programa

└── 📄 Simulador.xlsm       # Archivo principal

└── 📄 README.md           

# ⚡ Funcionalidades Implementadas
## ✅ Características Principales
### 1. 🔄 Ejecución de Código Ensamblador

- Carga de programas en lenguaje x86

- Ejecución paso a paso

- Visualización del estado en cada instrucción

### 2.👁️ Visualización de Componentes

- CPU: Unidad de Control, ALU y Registros

- Memoria: RAM, Caché y Memoria Virtual

- Pipeline: Flujo de instrucciones

- I/O: Unidad de Entrada/Salida

### 3.⚙️ Simulación de Políticas

- Gestión de memoria caché (LRU)

- Detección de riesgos en el pipeline

- Políticas de reemplazo

### 4.🔤 Traductor C a Ensamblado

- Conversión de código C simple a x86

- Ejecución del código traducido

## 🚀 Cómo Usar el Simulador
1. Abrir el archivo Excel que contiene el simulador

2. Cargar un programa en lenguaje ensamblador x86

3. Ejecutar paso a paso las instrucciones

4. Observar cómo se actualizan los componentes en cada paso

5. Analizar el comportamiento del pipeline y la memoria

## 👥 Equipo de Desarrollo
### Integrantes:
- Nicole Lozada Leon
  
- Dariana Pol Aramayo
  
- Krishna Ariany Lopez Melgar

### Documentación de Proyecto: 

[Google Docs](https://docs.google.com/document/d/12qxGSUhsfWce7u_80ljQAPUH7FlSxdMcGey3gcDZ-do/edit?usp=sharing)
### Presentación: 
[Canvas](https://www.canva.com/design/DAG7zXQHzOw/kvmCdCCMAJNdkTxkVYQPtw/edit?utm_content=DAG7zXQHzOw&utm_campaign=designshare&utm_medium=link2&utm_source=sharebutton)
### Gestion de tareas: 
[Kanban](https://github.com/users/arianylopez/projects/8)
