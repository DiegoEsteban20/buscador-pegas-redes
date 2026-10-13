# Buscador de empleos de Conectividad y Redes (PowerShell)

Busca ofertas de trabajo de redes, telecomunicaciones, soporte TI y ciberseguridad en varios portales chilenos, descarta lo que no es del área y arma un **Excel priorizado** que marca las ofertas **nuevas** y las más afines a tu perfil. Pensado para quien está terminando una carrera o práctica en el área y no quiere revisar cinco portales a mano todos los días.

Funciona en Windows PowerShell 5.1 o PowerShell 7.

---

## Tabla de contenidos
- [Para qué sirve](#para-qué-sirve)
- [Qué hace, paso a paso](#qué-hace-paso-a-paso)
- [Instalación](#instalación)
- [Uso](#uso)
- [El Excel que genera](#el-excel-que-genera)
- [Configuración (Config.psd1)](#configuración-configpsd1)
- [Dejarlo corriendo solo](#dejarlo-corriendo-solo)
- [Estructura del proyecto](#estructura-del-proyecto)
- [Preguntas frecuentes](#preguntas-frecuentes)
- [Conceptos aplicados](#conceptos-aplicados)
- [Aviso](#aviso)

---

## Para qué sirve

Buscar pega significa repetir las mismas búsquedas en GetOnBoard, Computrabajo, Chiletrabajos y más, filtrar el ruido ("community manager", "redes de agua potable", ventas) y acordarte de cuáles ya viste. Este script hace todo eso solo: consulta los portales, se queda solo con lo del área, ordena por lo más afín a tu perfil, elimina duplicados y te entrega un Excel listo, marcando lo nuevo desde la última vez.

---

## Qué hace, paso a paso

1. **Consulta los portales** con 15 búsquedas configurables (redes, telecomunicaciones, NOC, soporte TI, ciberseguridad, etc.):
   - **GetOnBoard** por su API pública (JSON).
   - **Computrabajo**, **Chiletrabajos** y **Talent.com** leyendo el HTML público.
2. **Filtra por el cargo** con expresiones regulares: deja los cargos del área y descarta falsos positivos como "redes sociales", "redes de apoyo", ventas, marketing o diseño.
3. **Filtra por nivel**: descarta Senior, Jefe, Líder, Gerente… (se puede apagar cuando tengas más experiencia).
4. **Filtra por ubicación**: solo Santiago / Región Metropolitana (o 100 % remoto), normalizando las tildes para no perder resultados.
5. **Calcula la relevancia** (★ a ★★★) según palabras clave y resalta prácticas y cargos junior.
6. **Elimina duplicados** entre portales y marca como **NUEVA** cada oferta que no habías visto.
7. **Genera un Excel** con varias hojas y te avisa con una notificación cuando hay ofertas nuevas.

---

## Instalación

No necesitas instalar nada aparte. La primera vez, el módulo `ImportExcel` (para crear el Excel) se instala solo.

Doble clic en `BuscarPegasAhora.cmd`.

---

## Uso

- **`BuscarPegasAhora.cmd`**: busca ahora y abre el Excel (tarda 2–4 minutos).
- **`InstalarAutomatico.cmd`**: deja la búsqueda programada cada vez que inicias sesión en Windows.

Por línea de comandos:

```powershell
.\BuscadorPegas.ps1 -SiempreAbrir   # busca ahora
.\BuscadorPegas.ps1 -Instalar       # automático al iniciar sesión
```

---

## El Excel que genera

| Hoja | Contenido |
| --- | --- |
| **Más afines** | Las ofertas más relacionadas con redes/conectividad (★★★ = muy afín) |
| **Nuevas** | Las que no habías visto antes |
| **Todas** | Todo lo que pasó los filtros |
| **Otros portales** | Links listos a LinkedIn, Laborum, Indeed, etc. (no permiten lectura automática) |

El nivel (práctica, junior, sin experiencia) va resaltado en amarillo, y la columna **Link** abre el aviso con un clic.

---

## Configuración (Config.psd1)

Ábrelo con el Bloc de notas. Lo principal:

| Opción | Qué controla |
| --- | --- |
| **Busquedas** | Los términos que se buscan en cada portal |
| **Incluir / Excluir** | Palabras del cargo que sí o que no te interesan (regex) |
| **ExcluirSenior** | `$true` descarta Senior/Jefe/Líder. Ponlo en `$false` cuando tengas más experiencia |
| **Portales** | Activa o desactiva cada portal |
| **SoloSantiago** | `$true` deja solo Santiago / RM y lo 100 % remoto. `$false` muestra todo Chile |
| **IncluirRemoto** | Incluye ofertas 100 % remotas aunque no digan Santiago |

---

## Dejarlo corriendo solo

`InstalarAutomatico.cmd` crea una tarea programada que corre al iniciar sesión (como máximo una vez cada 6 horas) y te avisa con una notificación cuando aparecen ofertas nuevas, sin que tengas que abrir nada.

---

## Estructura del proyecto

```
buscador-pegas-redes/
├─ BuscadorPegas.ps1         # script principal
├─ BuscarPegasAhora.cmd      # buscar ahora
├─ InstalarAutomatico.cmd    # programar al iniciar sesión
├─ Config.psd1               # búsquedas, filtros, ubicación, portales
├─ LEEME.txt                 # ayuda rápida en español
└─ README.md
```

---

## Preguntas frecuentes

**Me salen pocas ofertas.** Baja los filtros: pon `ExcluirSenior = $false` o `SoloSantiago = $false` en el Config.

**Un portal no trae nada.** Algunos cambian su HTML cada cierto tiempo; puedes desactivarlo en `Portales` y usar los links de la hoja "Otros portales".

**¿Es legal?** Solo lee páginas públicas, con pausas entre consultas para no sobrecargar los sitios.

---

## Conceptos aplicados

Consumo de API REST y scraping HTML responsable (pausas entre consultas, solo páginas públicas), expresiones regulares, normalización de texto (tildes), deduplicación, compatibilidad Windows PowerShell 5.1 / PowerShell 7, generación de Excel con `ImportExcel` y tareas programadas.

---

## Aviso

Herramienta de uso personal para buscar trabajo. Respeta los términos de cada portal y no la uses para consultas masivas.
