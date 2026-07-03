# Sistema de Formato de Revisión de Proyectos Eléctricos

Este proyecto es una aplicación web empresarial moderna construida en **Python** utilizando **Streamlit**, diseñada para automatizar la revisión técnica de proyectos eléctricos mediante un sistema de checklist interactivo. Reemplaza de manera eficiente una solución anterior basada en Google Sheets + Apps Script.

---

## Características Principales

1. **Búsqueda Dinámica de Proyectos**: Permite buscar por número de proyecto (**LCL**) y autocompletar automáticamente el Cliente, Distrito, Contratista y Supervisor asociados.
2. **Checklists Dinámicos por Tipo de Atención**: Adapta los requisitos de inspección en tiempo real según el tipo de atención seleccionado (Reforma, Reforma Sustancial, Factibilidad, etc.). Las preguntas están basadas en el checklist oficial de la empresa (39 preguntas).
3. **Cálculo de Estado en Tiempo Real**: Si existe al menos un requisito no conforme (`No`), el estado se calcula instantáneamente como `OBSERVADO`. De lo contrario, se mantiene como `CONFORME`.
4. **Generación de Reportes PDF Corporativos**: Compila un informe PDF profesional (con membrete, tablas estructuradas de datos generales, tabla detallada de checklist y sección de firmas) utilizando **ReportLab**.
5. **Historial Centralizado y Servidor de Descarga Virtual**: Registra las revisiones en una base de datos (con las respuestas en formato JSON) y permite reconstruir el PDF exacto de cualquier revisión histórica al instante, sin necesidad de almacenar archivos PDF físicos pesados en un servidor de archivos.
6. **Conectividad Híbrida**: Admite modo en la nube con **Google Sheets** (usando credenciales de cuenta de servicio en los secrets de Streamlit) y posee un **mecanismo de contingencia local** automático en Excel (`db_local.xlsx`) para poder ejecutarse sin conexión a Internet o sin configurar credenciales de Google Cloud.

---

## Estructura del Código

- **`app.py`**: Interfaz de usuario, diseño con columnas, CSS corporativo, sidebar de navegación y enrutamiento.
- **`checklists.py`**: Módulo maestro con las 39 preguntas oficiales y el mapeo correspondiente a los 6 tipos de atención soportados.
- **`database.py`**: Conectividad híbrida. Lee los datos generales del proyecto y escribe/lee el historial en Google Sheets o en el Excel local fallback.
- **`pdf_generator.py`**: Motor de dibujo PDF que diseña e imprime reportes con apariencia limpia e institucional.

---

## Requisitos de Instalación

1. **Python 3.12 o posterior**
2. Clonar o mover los archivos de este proyecto al directorio de tu preferencia.
3. Instalar las dependencias listadas en `requirements.txt`:

```bash
pip install -r requirements.txt
```

---

## Cómo Ejecutar la Aplicación

Para iniciar el servidor local de desarrollo de Streamlit, ejecuta en la terminal:

```bash
streamlit run app.py
```

La aplicación se abrirá automáticamente en tu navegador web predeterminado (usualmente en `http://localhost:8501`).

---

## Configuración de Google Sheets (En la Nube)

Para activar el almacenamiento en la nube en lugar del Excel local (`db_local.xlsx`):

1. Ve a la consola de Google Cloud e inicia un proyecto.
2. Habilita la API de Google Sheets y Google Drive.
3. Crea una cuenta de servicio de Google Cloud y genera una llave en formato JSON.
4. Comparte tu documento de Google Sheets (con acceso de edición) al correo de la cuenta de servicio (`tu-correo-servicio@...`).
5. Asegúrate de que tu Google Sheet contenga dos hojas llamadas:
   - **`Proyectos`**: Con las columnas `LCL`, `Cliente`, `Distrito`, `Contratista`, `Supervisor`.
   - **`Historial`**: Con las columnas `ID_Revision`, `Fecha`, `LCL`, `Cliente`, `Distrito`, `Contratista`, `Supervisor`, `Tipo_Atencion`, `Numero_Revision`, `Estado`, `Observaciones`, `Respuestas`.
6. En tu proyecto local, edita el archivo [`.streamlit/secrets.toml`](file:///C:/Users/P723919021/Documents/vscode/ICOS/.streamlit/secrets.toml), remueve los símbolos `#` de comentario y reemplaza los valores correspondientes con la información de tu cuenta de servicio y el enlace a tu Google Sheet.
