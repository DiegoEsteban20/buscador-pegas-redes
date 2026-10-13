@{
    # ===== Qué buscar (cada búsqueda se hace en cada portal activo) =====
    Busquedas = @(
        'tecnico en redes', 'redes y telecomunicaciones', 'conectividad', 'ingeniero en redes',
        'administrador de redes', 'telecomunicaciones', 'fibra optica', 'noc',
        'soporte ti', 'soporte tecnico', 'help desk', 'mesa de ayuda',
        'infraestructura ti', 'ciberseguridad', 'practica informatica'
    )

    # ===== Portales (pon $false para desactivar uno) =====
    Portales = @{
        GetOnBoard    = $true
        Computrabajo  = $true
        Chiletrabajos = $true
        Talent        = $true
    }

    # ===== Filtros sobre el nombre del cargo (expresiones regulares, sin importar mayúsculas) =====
    # El cargo debe tener al menos una de estas
    Incluir = 'redes|\bred\b|network|conectividad|telecom|fibra|\bnoc\b|\bsoc\b|soporte|support|help ?desk|mesa de ayuda|service ?desk|infraestructura|infrastructure|sysadmin|administrador de sistemas|ciberseguridad|cybersecurity|seguridad inform|seguridad de la informaci|firewall|fortinet|cisco|ccna|wi-?fi|wireless|cableado|data ?center|t[eé]cnico (inform|en computaci|ti\b|de ti\b|computacional)|\bti\b|\bit\b|cloud|plataforma|sd-?wan|vpn|switching|routing|telefon[ií]a|\bvoip\b'

    # Se descartan cargos con estas palabras (no son del área)
    Excluir = 'redes sociales|fibra de vidrio|legal|gestor.{0,3} de redes|administrativ|\baguas?\b|social media|community|marketing|contenido|audiovisual|dise[ñn]ador|ventas|vendedor|comercial|cobranza|\bagua\b|alcantarillado|sanitari|incendio|acu[ií]col|salm[oó]n|\bgas\b|el[eé]ctric|fr[ií]o|asistente social|trabajador social|redes de apoyo|enfermer|docente|profesor|chofer|conductor|bodeguero|cajero|contador|abogad|kinesi|psic[oó]log|nutricion|param[eé]dic|full.?stack|front.?end|back.?end|desarrollador|developer|programador'

    # Cargos de nivel alto (se descartan mientras ExcluirSenior = $true)
    ExcluirSenior = $true
    PalabrasSenior = '\bsenior\b|\bsr\.?\b|\bssr\b|semi.?senior|jefe|jefa|gerente|subgerente|\bhead\b|\blead\b|l[ií]der|arquitect|director|\bmanager\b|principal|coordinador'

    # Palabras que suman relevancia (aparecen primero en el Excel)
    MuyRelevante = 'redes|network|conectividad|telecom|fibra|\bnoc\b|cisco|fortinet|firewall|ccna|wi-?fi|wireless|cableado|switch|router|sd-?wan|vpn|data ?center|telefon[ií]a'
    Relevante    = 'soporte|support|help ?desk|mesa de ayuda|service ?desk|infraestructura|sysadmin|ciberseguridad|cybersecurity|\bsoc\b|\bti\b|\bit\b|t[eé]cnico|cloud'
    NivelEntrada = 'pr[aá]ctica|practicante|trainee|junior|\bjr\b|sin experiencia|reci[eé]n (egresad|titulad)|egresad|estudiante|ayudante|asistente|aprendiz'

    # ===== Ubicación =====
    # Solo ofertas en Santiago / Región Metropolitana (pon $false para ver todo Chile)
    SoloSantiago = $true
    # Incluir ofertas 100% remotas aunque no digan Santiago
    IncluirRemoto = $true
    # Comunas y palabras que cuentan como Santiago (sin tildes)
    Santiago = 'santiago|metropolitana|\br\.? ?m\b|providencia|las condes|vitacura|lo barnechea|nunoa|la reina|macul|penalolen|la florida|puente alto|san miguel|san joaquin|la cisterna|el bosque|la granja|la pintana|san ramon|lo espejo|pedro aguirre cerda|cerrillos|maipu|estacion central|quinta normal|lo prado|pudahuel|cerro navia|renca|quilicura|huechuraba|conchali|independencia|recoleta|san bernardo|buin|paine|calera de tango|colina|lampa|tiltil|padre hurtado|penaflor|talagante|isla de maipo|el monte|melipilla|curacavi|pirque|san jose de maipo'
    # Si el cargo menciona otra ciudad o región, se descarta
    OtrasCiudades = 'arica|iquique|tarapaca|antofagasta|calama|copiapo|atacama|la serena|coquimbo|ovalle|valparaiso|vina del mar|quillota|los andes|san antonio|rancagua|o.?higgins|talca|curico|maule|chillan|concepcion|biobio|bio-bio|temuco|araucania|valdivia|osorno|puerto montt|los lagos|chiloe|castro|coyhaique|punta arenas|magallanes|san pedro de atacama|mejillones|tocopilla'

    # ===== GetOnBoard =====
    NivelMaximo = 3      # 1 Sin experiencia, 2 Junior, 3 Semi Senior, 4 Senior, 5 Expert
    DiasMaximos = 30     # solo ofertas publicadas en los últimos N días
    Paises      = @('Chile', 'Remote')

    # ===== Computrabajo: solo ofertas publicadas en los últimos N días (1, 3, 7 o 15) =====
    ComputrabajoDias = 15
}
