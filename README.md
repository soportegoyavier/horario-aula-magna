# Horario Aula Magna

Sistema de gestión de horarios para el Aula Magna. Solo los editores autorizados ven el menú de administración.

---

## Instalación en Google Sheets

### Paso 1 — Abre el editor de Apps Script
En tu Google Sheet ve a **Extensiones → Apps Script**.

### Paso 2 — Pega el código principal
1. Borra todo lo que haya en `Code.gs`
2. Copia el contenido de `Code.gs` (de este repositorio) y pégalo
3. Guarda con `Ctrl+S`

### Paso 3 — Crea los archivos HTML
En el editor de Apps Script, para cada archivo HTML:
1. Clic en **"+"** (junto a "Archivos") → **HTML**
2. Nómbralo exactamente igual (sin `.html`):
   - `AddEvent`
   - `RemoveEvent`
   - `AllEvents`
3. Borra el contenido predeterminado y pega el contenido de cada archivo `.html` de este repositorio

### Paso 4 — Guarda y recarga
1. Guarda todos los archivos (`Ctrl+S`)
2. Cierra y vuelve a abrir el Google Sheet
3. Verás el menú **"Aula Magna"** en la barra superior *(solo si tienes permisos de editor)*

---

## Estructura de la hoja

Cada pestaña del Sheet corresponde a un mes (MAYO, JUNIO, etc.)  
Cada semana sigue este formato:

```
Fila X:    [ Semana del DD al DD ] (celdas combinadas)
Fila X+1:  [ HORA | Lunes | Martes | Miércoles | Jueves | Viernes | Sabado | Domingo ]
Fila X+2:  [ 07:00am | ... ]
Fila X+3:  [ 08:00am | ... ]
  ...
Fila X+15: [ 08:00pm | ... ]
Fila X+16: (fila vacía separadora)
```

---

## Categorías y colores

| Categoría           | Color      |
|---------------------|------------|
| Ensayo              | Azul claro |
| Evento/Conferencia  | Naranja    |
| Ceremonia           | Lila       |
| Clase/Taller        | Verde menta|

---

## Control de acceso

- **Editores** (agregados en "Compartir"): ven el menú Aula Magna y pueden agregar/eliminar eventos
- **Lectores**: solo ven el horario, sin menú ni opciones de edición

---

## Archivos del proyecto

| Archivo         | Descripción                                      |
|-----------------|--------------------------------------------------|
| `Code.gs`       | Lógica principal: menú, agregar, eliminar eventos |
| `AddEvent.html` | Sidebar para agregar un nuevo evento             |
| `RemoveEvent.html` | Sidebar para eliminar eventos existentes      |
| `AllEvents.html`| Sidebar con listado completo de eventos          |
