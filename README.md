# El duendecito de Vianni

Aplicacion local de Windows para preparar automaticamente documentos de nuevos empleados a partir del archivo exportado por Mercury Soluciones.

## Que hace

La aplicacion vigila la carpeta Descargas y busca estos archivos:

- `GridViewExports.xls`
- `GridViewExport.xls`
- `GridViewExports.xlsx`
- `GridViewExport.xlsx`

Cuando encuentra uno, lo copia a `imported_files`, lo renombra como `nuevasEntradas_dd.mm.yyyy`, lee los empleados, genera documentos desde las plantillas y guarda todo en una carpeta fechada dentro de `output`.

## Instalacion para desarrollo

Requisitos:

- Windows 10/11
- Python 3.12
- Microsoft Office instalado si se necesita imprimir o trabajar con plantillas antiguas `.doc` y `.xls`

Pasos:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m el_duendecito_de_vianni.app
```

## Estructura

```text
plantillas/
  Brothers/
  Ines/
imported_files/
output/
logs/
config.json
politica_horario.xlsx
demo/
src/
tests/
```

## Configuracion

El archivo `config.json` se crea automaticamente la primera vez que se ejecuta la aplicacion. Tambien puede copiarse desde `config.example.json`.

Desde la ventana de Configuracion se puede cambiar:

- Carpeta de Descargas
- Carpeta de plantillas
- Carpeta de salida
- Tabla de horarios
- Direccion, usuario y modo visible/invisible para Mercury
- Intervalo de escaneo
- Confirmacion antes de borrar el archivo descargado
- Impresora
- Inicio minimizado en bandeja
- Inicio automatico con Windows

La contrasena de Mercury no se guarda en `config.json`. Se guarda en el Administrador de credenciales de Windows y se puede reemplazar desde Configuracion escribiendo una nueva contrasena, sin pedir la anterior.

## Mercury

La integracion de Mercury permite configurar la direccion, el usuario, la compania, el reporte, la contrasena y si Playwright debe abrir el navegador visible o invisible. El boton `Mercury` entra al sitio, abre el generador de reportes, descarga el reporte configurado y luego procesa el archivo descargado con el mismo flujo normal de documentos.

Las fotos de empleados quedan para una fase posterior.

## Plantillas

Coloque las plantillas en una de estas carpetas:

- `plantillas/Brothers`
- `plantillas/Ines`

Cuando el campo `Compania` del exporte contiene exactamente `Supermercado Ines`, la aplicacion usa solo las plantillas de `plantillas/Ines`. Para cualquier otro valor de `Compania`, usa solo las plantillas de `plantillas/Brothers`.

Dentro de cada carpeta, las plantillas se procesan en orden alfabetico, por eso se recomienda nombrarlas asi:

- `01_Contrato.docx`
- `02_Formulario.xlsx`
- `10_Declaracion.docx`

Las plantillas que solo deben copiarse e imprimirse, sin reemplazar marcadores, pueden incluir `__STATIC__` en el nombre. Por ejemplo: `03__STATIC__Reglamento.docx`.

Los marcadores usan este formato:

```text
{{Nombre Empleado}}
{{Cedula}}
{{Posición}}
{{Salario Base}}
```

El texto dentro de `{{ }}` debe coincidir exactamente con el encabezado de la columna en Excel, incluyendo espacios, mayusculas y acentos.

Tambien puede aplicar formatos simples agregando `|formato` al marcador:

```text
{{Salario Base|money}}      -> 18,800.00
{{Numero|int}}              -> 295
{{Fecha Ingreso|date}}      -> 04/07/2026
{{Sexo|tratamiento}}        -> Sra. / Sr.
Estimad{{Sexo|genero}}      -> Estimada / Estimado
Estimad{{Sexo|genero_plural}} -> Estimadas / Estimados
colaborador{{Sexo|genero_sustantivo}} -> colaboradora / colaborador
colaborador{{Sexo|genero_sustantivo_plural}} -> colaboradoras / colaboradores
```

Los formatos disponibles son `money`, `int`, `date`, `tratamiento`, `genero`, `genero_plural`, `genero_sustantivo` y `genero_sustantivo_plural`. Para `tratamiento`, los valores femeninos como `F`, `Fem`, `Femenino` o `Mujer` generan `Sra.`, y los valores masculinos como `M`, `Masc`, `Masculino` o `Hombre` generan `Sr.`. `genero` genera `a` u `o`, y `genero_plural` genera `as` u `os`. `genero_sustantivo` genera `a` o vacio, y `genero_sustantivo_plural` genera `as` o `es`. Si no se indica formato, el valor se inserta como texto normal.

### Horario laboral

Para construir la frase de politica de horario, coloque la tabla configurable en `politica_horario.xlsx` o seleccione otro archivo desde Configuracion. La aplicacion busca el valor del campo `Política Horario` en la columna `horario1 (de GridViewExport.xls)` y genera el campo:

```text
{{Horario Laboral}}
```

La tabla debe incluir las columnas `horario1 (de GridViewExport.xls)`, `horario2`, `dias`, `break` y `feriados`. Si `dias` es `5`, se genera una frase de lunes a viernes y sabado. Si `dias` es `6`, se genera una frase de lunes a sabado, domingo y un dia libre semanal. Cuando `feriados` es `1`, se agrega que el horario incluye dias feriados; cuando es `0`, esa parte se omite.

## Formatos soportados

- Exportes de empleados: `.xlsx` y `.xls`
- Plantillas modernas: `.docx` y `.xlsx`
- Plantillas antiguas: `.doc` y `.xls`

La primera version procesa directamente `.docx` y `.xlsx`. Las plantillas antiguas `.doc` y `.xls` se detectan y se registran como omitidas para evitar generar documentos corruptos. Para usarlas, conviertalas a `.docx` o `.xlsx`. En Windows con Microsoft Office instalado se puede extender el modulo `office.py` para convertirlas mediante automatizacion COM.

## Bandeja del sistema

Al iniciar, la aplicacion queda minimizada en la bandeja del sistema. El boton X de la ventana solo la oculta; para cerrar completamente use `Salir` desde el menu de la bandeja.

El menu de la bandeja incluye:

- Abrir El duendecito de Vianni
- Procesar ahora
- Iniciar monitoreo
- Detener monitoreo
- Abrir carpeta de salida
- Configuracion
- Ver registro
- Salir

## Empaquetar como EXE

Ejecute:

```powershell
.\build_exe.ps1
```

El ejecutable se genera en:

```text
dist\ElDuendecitoDeVianni.exe
```

No ejecute el archivo que aparece dentro de `build\`. Esa carpeta contiene archivos temporales de PyInstaller.

## Pruebas

```powershell
.\.venv\Scripts\python.exe -m pytest
```

Las pruebas cubren reemplazo en Word, Excel, campos con acentos, celdas vacias, marcadores faltantes, deteccion de duplicados, nombres invalidos y orden de impresion por nombre de archivo.

## Limitaciones conocidas

- La impresion automatica depende de Windows y de las aplicaciones asociadas a cada tipo de archivo.
- La seleccion fina de impresora queda preparada en configuracion, pero Windows puede usar la impresora predeterminada segun la aplicacion que imprima el documento.
- Las plantillas `.doc` y `.xls` antiguas se omiten en esta primera version para evitar corrupcion. Se recomienda convertirlas a `.docx` y `.xlsx`.
- Los cuadros de texto complejos en Word pueden requerir conversion o ajuste manual si `python-docx` no puede acceder a su contenido.
