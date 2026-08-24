# Importador de contratos Gana Energia

Proyecto en JavaScript con Vite + React para:

1. Cargar un Excel con contratos.
2. Validar que la primera hoja use los encabezados esperados.
3. Generar un `.txt` con un script listo para pegar en la consola del navegador dentro de `https://externos.ganaenergia.com`.
4. Ejecutar los contratos desde la propia web de Gana Energia en dos modos:
   - todas las filas del Excel
   - solo las filas cuya `fecha_creacion` coincida con la fecha local del sistema

## Requisitos

- Node.js 18 o superior
- npm

## Como ejecutarlo

Instala dependencias:

```bash
npm install
```

Arranca el proyecto:

```bash
npm run dev
```

Despues abre en el navegador la URL que muestre Vite, normalmente:

```bash
http://localhost:5173
```

## Como usarlo

### 1) Descargar plantilla

Pulsa `Descargar plantilla ejemplo` para generar un Excel con la estructura esperada.

La primera hoja debe contener exactamente los encabezados de la plantilla.

### 2) Cargar Excel

Sube el archivo `.xlsx` o `.xls`.

La app:

- lee la primera hoja
- valida cabeceras
- cuenta todas las filas
- detecta cuantas filas tienen `fecha_creacion` igual a la fecha local del sistema en formato `YYYY-MM-DD`

### 3) Generar TXT para consola

La app ofrece dos descargas:

- `Descargar TXT con todas las filas`
- `Descargar TXT filtrado por fecha_creacion`

Cada `.txt` incluye:

- los payloads ya convertidos a JSON anidado
- la logica para hacer `POST /contract`
- deteccion automatica de token desde `localStorage` o `sessionStorage`
- fallback para pegar un bearer token manualmente si hiciera falta
- descarga automatica de un JSON con el resultado final

### 4) Ejecutar el TXT en la web

1. Inicia sesion en `https://externos.ganaenergia.com`.
2. Abre cualquier pagina de la web donde la sesion siga activa.
3. Abre la consola del navegador.
4. Copia y pega el contenido completo del `.txt`.
5. Pulsa Enter.

El script enviara un `POST` por cada contrato contra:

```text
https://externos.ganaenergia.com/contract
```

Al terminar:

- mostrara los resultados en consola
- dejara los datos en `window.__ganaenergiaLastResults`
- descargara un archivo Excel con el resumen, incluyendo el body de respuesta por fila

## Build de produccion

```bash
npm run build
```

## App de escritorio

Modo escritorio en desarrollo:

```bash
npm run desktop:dev
```

Generar app para macOS:

```bash
npm run dist:mac
```

Generar instalador para Windows:

```bash
npm run dist:win
```

El resultado se genera en `release/`.
