# Buscador de empleos de Conectividad y Redes (PowerShell)

Busca ofertas de trabajo de redes, telecomunicaciones, soporte TI y ciberseguridad en varios portales chilenos, descarta lo que no es del área y arma un **Excel priorizado** con las ofertas nuevas.

## Qué hace

- Consulta **GetOnBoard** (API pública JSON), **Computrabajo**, **Chiletrabajos** y **Talent.com** (HTML) con 15 búsquedas configurables.
- Filtra por expresiones regulares: incluye cargos del área y descarta falsos positivos ("redes sociales", "redes de agua", ventas…).
- Filtra por **nivel** (excluye Senior/Jefe/Líder) y por **ubicación** (Santiago / Región Metropolitana o remoto), normalizando tildes.
- Calcula una **relevancia** (★ a ★★★) según palabras clave y destaca prácticas y cargos junior.
- Elimina duplicados entre portales y marca como **NUEVA** cada oferta no vista antes.
- Genera un Excel con hojas *Más afines*, *Nuevas*, *Todas* y links listos a portales que no permiten lectura automática.
- Se programa al iniciar sesión en Windows y avisa con una notificación cuando hay ofertas nuevas.

## Uso

```powershell
.\BuscadorPegas.ps1 -SiempreAbrir   # busca ahora
.\BuscadorPegas.ps1 -Instalar       # automático al iniciar sesión
```

Toda la configuración (búsquedas, filtros, ubicación, portales) está en `Config.psd1`.

## Conceptos aplicados

Consumo de API REST y scraping HTML responsable (pausas entre consultas, solo páginas públicas), expresiones regulares, deduplicación, compatibilidad Windows PowerShell 5.1 / PowerShell 7, generación de Excel con `ImportExcel` y tareas programadas.
