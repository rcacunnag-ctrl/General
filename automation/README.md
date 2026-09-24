# Registro automático de avance – Proyecto Jupiter

Macro VBA que guarda una **foto diaria** y otra **semanal** del avance del cronograma en el libro `Property_Scope_Plan_Jupiter_v2.xlsx`. Esas fotos alimentan la curva S del Dashboard y el gráfico de tendencia diaria.

## Qué hace

| Momento | Qué registra |
|---|---|
| Al **abrir** el libro | La foto del día, con el estado tal como quedó en el último guardado |
| Al **guardar** | Actualiza la foto del día con los últimos cambios |
| Botón **Record progress now** (Dashboard) | Registra en ese momento y confirma con un mensaje |

- **Progress Log** (diario): una fila por día. Si se registra de nuevo el mismo día, la fila se sobrescribe; no se duplica.
- **Weekly Log** (semanal): una fila por semana, con cierre el viernes. Cada registro actualiza la semana en curso ("Week-to-date"). Pasado el viernes, la fila queda "Final". Lo que se registra el sábado o el domingo cuenta para el viernes anterior.
- **Contenido de cada foto:**
  - % planeado y % real, desvío, SPI y desvío en horas.
  - Actividades atrasadas y actividades de las próximas 2 semanas.
  - Estado de salud del proyecto.
  - Hitos cumplidos, hitos atrasados y el primer hito atrasado.
  - Planeado y real de las 11 secciones.
- **La foto siempre es "a hoy".** La macro pone temporalmente la fecha de corte del Dashboard (C4) en la fecha del día, copia los valores y restaura lo que había en C4, ya sea una fórmula o una fecha fija.
- **Los valores quedan congelados**, así que el historial no cambia si después se modifica el cronograma.

## Instalación (Windows + Excel 2016/2019/2021/365)

1. Pon en **una misma carpeta** el libro `Property_Scope_Plan_Jupiter_v2.xlsx` y los archivos de `automation/`. También funciona con el libro en la carpeta superior.
2. Cierra Excel y haz **doble clic en `Install.cmd`**.
   - El instalador desbloquea los archivos descargados, busca el libro solo y genera `Property_Scope_Plan_Jupiter_v2.xlsm` en la misma carpeta. El `.xlsx` no se modifica.
   - Activa solo durante la instalación el permiso "Trust access to the VBA project object model" y al final lo deja como estaba.
   - Si algo falla, la ventana muestra `FAILED at step: …` con el motivo. Copia ese texto para diagnosticarlo.
3. Abre el `.xlsm` y pulsa **Enable Content** la primera vez.
4. Desde ahí, **trabaja siempre en el `.xlsm`**.

Si prefieres la consola, desde la carpeta de los archivos:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Install-ProgressLogger.ps1 -EnableVbomAccess
# o indicando el libro:
powershell -NoProfile -ExecutionPolicy Bypass -File .\Install-ProgressLogger.ps1 -Workbook "C:\ruta\Property_Scope_Plan_Jupiter_v2.xlsx" -EnableVbomAccess
```

**Errores comunes**
| Mensaje | Causa y solución |
|---|---|
| `Could not find Property_Scope_Plan_Jupiter*.xlsx` | El libro no está junto al instalador. Muévelo o usa `-Workbook`. |
| `IT policy blocks programmatic access` | TI bloquea el acceso al proyecto VBA. Usa la instalación manual. |
| `Excel refused access to the VBA project` | Actívalo en File > Options > Trust Center > Trust Center Settings > Macro Settings > "Trust access to the VBA project object model" y repite. |
| `The output file is open` | Cierra el `.xlsm` en Excel y repite. |
| Scripts deshabilitados en el sistema | Usa `Install.cmd`, que ya ejecuta con `-ExecutionPolicy Bypass`. |

### Instalación manual (Mac, o si el script está bloqueado)
1. Abre el `.xlsx` y pulsa `Alt+F11` (en Mac: Tools > Macro > Visual Basic Editor).
2. Importa el módulo con **File > Import File…** > `ProgressLogger.bas`.
3. Abre el módulo **ThisWorkbook** (en Excel en español se llama **EsteLibro**) y pega el contenido de `ThisWorkbook.txt`.
4. Opcional, para agregar el botón: en el Dashboard ve a Developer > Insert > Button, asígnale la macro `RecordProgressNow` y ponle el texto "Record progress now".
5. Guarda como **Excel Macro-Enabled Workbook (.xlsm)**.

## Uso diario

1. En **Scope Checklist**, actualiza el estado (columna E) y el % de avance (columna I). Una actividad Complete cuenta como 100%.
2. Guarda. La foto del día se actualiza sola.
3. Revisa el **Dashboard**. En la celda H9 aparece la fecha y hora de la última foto registrada.

## Limitaciones

- **Solo registra cuando alguien abre o guarda el libro.** Si nadie lo abre un día, ese día no queda en el historial: el avance real de fechas pasadas no se puede reconstruir después. Basta con abrirlo una vez por semana para no perder la semana.
- **Semanas antes del arranque o después de las 30 del plan:** la macro las agrega al final del Weekly Log. La curva S del Dashboard cubre las 30 semanas planeadas (hasta el 30-abr-2027).
- **Si abrir el libro produce cambios**, Excel preguntará si quieres guardar al cerrar. Es normal: es la foto del día.
- **Excel en la web (OneDrive/SharePoint) no ejecuta VBA.** Si el libro se mueve a Microsoft 365, habrá que migrar a Office Script + Power Automate; la lógica es la misma.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `ProgressLogger.bas` | Módulo VBA con el registro diario y semanal |
| `ThisWorkbook.txt` | Eventos que disparan el registro (`Workbook_Open`, `Workbook_BeforeSave`) |
| `Install.cmd` | Lanzador de doble clic del instalador |
| `Install-ProgressLogger.ps1` | Instalador que genera el `.xlsm` con la macro, los eventos y el botón |
