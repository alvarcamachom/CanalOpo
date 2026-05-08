const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, Footer, PageBreak
} = require('docx');
const fs = require('fs');

const AZUL = "1F3864";
const AZUL2 = "2E75B6";
const AZUL3 = "D6E4F0";
const AZUL4 = "EBF3FB";
const VERDE = "1E5631";
const VERDE2 = "2D7D46";
const VERDE3 = "D5E8D4";
const NARANJA = "7B3F00";
const NARANJA2 = "C06000";
const NARANJA3 = "FFE6CC";
const ROJO = "641E16";
const ROJO2 = "C0392B";
const ROJO3 = "FADBD8";
const GRIS = "5D6D7E";
const GRIS2 = "D5D8DC";
const GRIS3 = "F2F3F4";
const BLANCO = "FFFFFF";

const brd = (c="CCCCCC",s=4) => ({style:BorderStyle.SINGLE,size:s,color:c});
const bords = (c="CCCCCC",s=4) => ({top:brd(c,s),bottom:brd(c,s),left:brd(c,s),right:brd(c,s)});

const cell = (text, {w,fill,tc=AZUL,bold=false,sz=18,al=AlignmentType.LEFT,cs=1,it=false}={}) =>
  new TableCell({
    borders: bords(fill||"CCCCCC"),
    width:{size:w,type:WidthType.DXA},
    shading: fill?{fill,type:ShadingType.CLEAR}:undefined,
    margins:{top:80,bottom:80,left:120,right:120},
    columnSpan:cs,
    children:[new Paragraph({alignment:al,children:[new TextRun({text:text??"",bold,italic:it,size:sz,color:tc,font:"Arial"})]})]
  });

// Cabecera de tabla de artículos
const artRow = (num, contenido, fillRow) =>
  new TableRow({ children: [
    cell(num, {w:1440,fill:fillRow||AZUL3,bold:true,sz:18,tc:AZUL}),
    cell(contenido, {w:8160,fill:BLANCO,sz:18,tc:"000000"})
  ]});

// Fila de sección dentro de tabla
const secRow = (texto, fill) =>
  new TableRow({ children: [
    cell(texto, {w:9600,fill,bold:true,sz:19,tc:BLANCO,cs:2})
  ]});

const h1 = (text,fill=AZUL,tc=BLANCO) => new Paragraph({
  heading:HeadingLevel.HEADING_1,
  children:[new TextRun({text,bold:true,size:28,color:tc,font:"Arial"})],
  shading:{fill,type:ShadingType.CLEAR},
  spacing:{before:400,after:200},indent:{left:200,right:200}
});

const h2 = (text,fill=AZUL2,tc=BLANCO) => new Paragraph({
  heading:HeadingLevel.HEADING_2,
  children:[new TextRun({text,bold:true,size:24,color:tc,font:"Arial"})],
  shading:{fill,type:ShadingType.CLEAR},
  spacing:{before:280,after:100},indent:{left:200}
});

const h3 = (text,tc=AZUL) => new Paragraph({
  heading:HeadingLevel.HEADING_3,
  children:[new TextRun({text,bold:true,size:21,color:tc,font:"Arial"})],
  spacing:{before:180,after:80},
  border:{bottom:{style:BorderStyle.SINGLE,size:2,color:AZUL3}}
});

const body = (text,bold=false,indent=200) => new Paragraph({
  children:[new TextRun({text,size:20,bold,font:"Arial"})],
  spacing:{before:60,after:60},indent:{left:indent}
});

const art = (num,text) => new Paragraph({
  children:[
    new TextRun({text:`Art. ${num}.  `,bold:true,size:20,color:AZUL2,font:"Arial"}),
    new TextRun({text,size:20,font:"Arial"})
  ],
  spacing:{before:80,after:80},indent:{left:360}
});

const bullet = (text) => new Paragraph({
  numbering:{reference:"bul",level:0},
  children:[new TextRun({text,size:20,font:"Arial"})],
  spacing:{before:40,after:40}
});

const sep = () => new Paragraph({children:[new TextRun("")],spacing:{before:160,after:160},border:{bottom:{style:BorderStyle.SINGLE,size:4,color:AZUL3}}});
const esp = () => new Paragraph({children:[new TextRun("")],spacing:{before:80,after:80}});

// Tabla simple de 2 col: etiqueta / valor
const metaRow = (label,value,even) =>
  new TableRow({children:[
    cell(label,{w:2800,fill:AZUL3,bold:true,sz:18,tc:AZUL}),
    cell(value,{w:6800,fill:even?GRIS3:BLANCO,sz:18,tc:"000000"})
  ]});

const metaTable = (rows) => new Table({
  width:{size:9600,type:WidthType.DXA},
  columnWidths:[2800,6800],
  rows: rows.map(([l,v],i)=>metaRow(l,v,i%2===1))
});

const doc = new Document({
  numbering:{config:[
    {reference:"bul",levels:[{level:0,format:LevelFormat.BULLET,text:"\u2022",alignment:AlignmentType.LEFT,
      style:{paragraph:{indent:{left:720,hanging:360}}}}]}
  ]},
  styles:{
    default:{document:{run:{font:"Arial",size:20}}},
    paragraphStyles:[
      {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,
        run:{size:28,bold:true,font:"Arial",color:BLANCO},paragraph:{spacing:{before:400,after:200},outlineLevel:0}},
      {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,
        run:{size:24,bold:true,font:"Arial",color:BLANCO},paragraph:{spacing:{before:280,after:100},outlineLevel:1}},
      {id:"Heading3",name:"Heading 3",basedOn:"Normal",next:"Normal",quickFormat:true,
        run:{size:21,bold:true,font:"Arial",color:AZUL},paragraph:{spacing:{before:180,after:80},outlineLevel:2}}
    ]
  },
  sections:[{
    properties:{page:{size:{width:11906,height:16838},margin:{top:1260,right:1080,bottom:1260,left:1080}}},
    footers:{default:new Footer({children:[new Paragraph({
      alignment:AlignmentType.CENTER,
      border:{top:{style:BorderStyle.SINGLE,size:4,color:AZUL3}},
      children:[
        new TextRun({text:"III Convenio Colectivo — Canal de Isabel II, S.A.M.P. (2025-2028)   |   Pág. ",size:16,font:"Arial",color:"888888"}),
        new TextRun({children:[PageNumber.CURRENT],size:16,font:"Arial",color:"888888"})
      ]
    })]})},
    children:[

      // ═══════════════════════════
      // PORTADA
      // ═══════════════════════════
      new Paragraph({alignment:AlignmentType.CENTER,shading:{fill:AZUL,type:ShadingType.CLEAR},
        children:[new TextRun({text:"III CONVENIO COLECTIVO",bold:true,size:48,color:BLANCO,font:"Arial"})],
        spacing:{before:600,after:0},indent:{left:300,right:300}}),
      new Paragraph({alignment:AlignmentType.CENTER,shading:{fill:AZUL2,type:ShadingType.CLEAR},
        children:[new TextRun({text:"Canal de Isabel II, S.A.M.P.",bold:false,size:32,color:BLANCO,font:"Arial"})],
        spacing:{before:0,after:0},indent:{left:300,right:300}}),
      new Paragraph({alignment:AlignmentType.CENTER,shading:{fill:AZUL2,type:ShadingType.CLEAR},
        children:[new TextRun({text:"Vigencia: 1 de enero de 2025 – 31 de diciembre de 2028",bold:false,size:22,color:BLANCO,font:"Arial"})],
        spacing:{before:0,after:400},indent:{left:300,right:300}}),
      esp(),
      metaTable([
        ["Empresa","Canal de Isabel II, S.A.M.P. — NIF A86488087"],
        ["Domicilio social","C/ Santa Engracia, 125, 28003 Madrid"],
        ["Registro Mercantil","Madrid, Tomo 29.733, Folio 86, Sección 8, Hoja M-534929"],
        ["Vigencia","1 de enero de 2025 hasta el 31 de diciembre de 2028"],
        ["Prórroga","Año a año si ninguna parte lo denuncia con 3 meses de antelación"],
        ["Ámbito personal","Toda la plantilla con cualquier modalidad de contratación (excluidos directivos y alta dirección)"],
        ["Ámbito territorial","Instalaciones de la Comunidad de Madrid, embalse El Vado (Castilla-La Mancha) y embalse La Aceña (Castilla y León)"],
        ["Número de títulos","IX Títulos, 106 artículos, 24 Disposiciones Adicionales, 4 Transitorias, 1 Final y 9 Anexos"],
      ]),
      esp(),
      new Paragraph({children:[new PageBreak()]}),

      // ═══════════════════════════
      // TÍTULO I — DISPOSICIONES GENERALES
      // ═══════════════════════════
      h1("TÍTULO I — DISPOSICIONES GENERALES"),
      esp(),
      art("1","Ámbito funcional: regula las relaciones laborales de Canal de Isabel II, S.A.M.P., cuya actividad comprende el ciclo integral del agua y sus servicios auxiliares."),
      art("2","Ámbito territorial: aplicable a las instalaciones de la Comunidad de Madrid y a los embalses de El Vado (Castilla-La Mancha) y La Aceña (Castilla y León)."),
      art("3","Ámbito personal: aplicable a toda la plantilla en activo en el territorio señalado, con independencia de la modalidad de contratación. Quedan excluidos quienes estén vinculados por contrato de alta dirección y las personas trabajadoras que desempeñen puestos de director/a."),
      art("4","Ámbito temporal: vigencia desde el 1 de enero de 2025 hasta el 31 de diciembre de 2028. Se prorroga automáticamente año a año salvo denuncia expresa con un mínimo de 3 meses de antelación."),
      art("5","Denuncia del convenio: mientras no se alcance un nuevo acuerdo, se mantiene en vigor la totalidad del contenido."),
      art("6","Vinculación a la totalidad: los acuerdos forman un todo orgánico e indivisible. Si se declarara nula alguna cláusula judicialmente, se negociará de nuevo preservando el equilibrio general."),
      art("7","Comisión paritaria: 7 representantes de la dirección + 7 de las secciones sindicales con presencia en el comité de empresa (al menos 2 de cada parte que hayan formado parte de la comisión negociadora). Funciones: interpretación del convenio, vigilancia del cumplimiento, intervención en conflictos colectivos, desarrollo de aspectos habilitados, e inaplicación de condiciones de trabajo."),
      art("8","Resolución de conflictos: con carácter previo a cualquier conflicto colectivo jurídico, las partes deben someter la cuestión a la Comisión Paritaria (plazo: 15 días naturales). Si no hay acuerdo, pueden acudir al Servicio Interconfederal de Mediación y Arbitraje (SIMA)."),
      art("9","Régimen de incompatibilidades: el personal está sujeto a la Ley 53/1984 de incompatibilidades del personal al servicio de las AA.PP. Toda actividad adicional (pública o privada) requiere previo reconocimiento de compatibilidad."),
      art("10","Confidencialidad: obligación de no divulgar información de la empresa ni utilizarla en interés propio o de terceros. Esta obligación se mantiene también tras la extinción del contrato."),
      art("11","Acoso en el trabajo: tolerancia cero frente a cualquier tipo de acoso. Se aplican los principios de la Ley 4/2023 LGTBI y las Leyes de la Comunidad de Madrid 5/2005 y 3/2016. El protocolo de actuación figura en el Anexo VIII."),
      art("12","Compensación y absorción: las condiciones del convenio forman un todo. Los beneficios individuales superiores se respetan con carácter estrictamente personal. Las garantías del CGI (Compromiso de Garantías Individuales) se rigen por la Disposición Adicional Primera."),
      sep(),

      // ═══════════════════════════
      // TÍTULO II — ORGANIZACIÓN DEL TRABAJO
      // ═══════════════════════════
      h1("TÍTULO II — ORGANIZACIÓN DEL TRABAJO"),

      h2("Capítulo I — Organización del trabajo"),
      art("13","Competencia: la facultad de organización corresponde a la Dirección de la Empresa, respetando la legislación vigente y el convenio."),
      art("14","Organización del trabajo: objetivo de alcanzar el óptimo de productividad con la mejor utilización de recursos. La dirección informará trimestralmente de las previsiones de subcontratación y establecerá fórmulas de participación (reuniones de grupo, equipos de enriquecimiento de tareas)."),
      art("15","Plantilla: instrumento organizativo sujeto a la Orden de 16 de enero de 2012. Las modificaciones se someten a informe de la representación de los trabajadores en un plazo de 5 días. La dirección entrega semestralmente al comité y secciones sindicales copia digital de la plantilla con denominación del puesto, ubicación, jornada y nivel retributivo. Orden de preferencia para cubrir vacantes: 1.º acoplamiento médico, 2.º reasignación interna, 3.º movilidad geográfica, 4.º excedencias pendientes de reingreso."),

      h2("Capítulo II — Movilidad"),
      art("16","Principio general de movilidad: facultad organizativa para cambiar, permutar, trasladar y ajustar personas a puestos, con los límites de la Ley y el Convenio."),
      h3("Sección 1ª — Movilidad Funcional"),
      art("17","Principio general: movilidad dentro del grupo profesional y área funcional conforme a las titulaciones requeridas y respetando la dignidad del trabajador."),
      art("18","Trabajo de inferior y superior grupo profesional: el desempeño de funciones de superior grupo no puede exceder de 6 meses (salvo ausencia temporal del titular). El desempeño de funciones de inferior grupo no excederá de 15 días ni generará consolidación. La Dirección informará previamente al Comité de Empresa en todos los supuestos."),
      art("19","Polivalencia: a efectos de correcto aprovechamiento de la jornada se establece la tabla de polivalencias del Anexo V. Se facilitará la formación necesaria para ejercerla."),
      h3("Sección 2ª — Movilidad organizativa y geográfica"),
      art("20","Principios: la Dirección determina las razones técnicas, organizativas y productivas. La movilidad geográfica no puede utilizarse con fines discriminatorios ni afectar a representantes legales salvo razones objetivas. Se mantienen las retribuciones fijas; las variables serán las del nuevo destino."),
      art("21","Permutas: entre trabajadores del mismo grupo, área funcional e igual denominación de puesto, en diferentes servicios/localidades. La dirección resuelve en 2 meses; si no hay resolución, se entiende denegada."),
      art("22","Desplazamientos temporales: hasta 2 semanas: compensación de kilómetros. Entre 2 semanas y 3 meses (dentro del territorio del convenio): 3 días de descanso si el destino dista menos de 30 km del domicilio habitual; 6 días si supera los 30 km. Se efectúan de forma rotatoria. Fuera del ámbito territorial: voluntarios, máximo 12 meses, con permisos específicos por duración."),
      art("23","Traslados: el traslado a petición del trabajador no genera compensación. Traslados forzosos: se efectúan siguiendo criterios de proximidad domicilio/nuevo destino, menor antigüedad y menor edad. Indemnizaciones por traslado forzoso a municipio diferente en función de la distancia (de 15 km a 19,9 km: 1.378,24 €; de 20 a 29,9 km: 3.097,76 €; de 30 a 40 km: 4.817,28 €; a partir de 40 km, 1.719,52 € por cada 10 km adicionales). Alternativa: percepción mensual por kilometraje diario. Traslados a más de 40 km: solo posibles dos veces por trabajador; el trabajador puede rescindir con indemnización de 2 mensualidades/año de antigüedad."),

      h2("Capítulo III — Ordenación del tiempo de trabajo"),
      art("24","Jornada de trabajo: 1.562 horas anuales (mientras la Comunidad de Madrid autorice la reducción; hasta entonces se mantienen las 1.673,5 horas anuales vigentes por la D.A. 144 de la Ley 6/2018). Máximo 9 horas diarias de trabajo efectivo. El exceso o defecto respecto a la jornada general se compensa con descanso equivalente en los 6 meses siguientes. El teletrabajo se regula en el Anexo IX."),
      art("25","Jornada en régimen de turnos: aplicable a los puestos con turnos descritos en el art. 76.1. Las horas nocturnas en turno de noche generan una bolsa del 2% en horas de descanso a disfrutar el año siguiente."),
      art("26","Calendarios laborales: se elaboran antes del 20 de octubre de cada año, con publicidad mínima de 10 días naturales y aprobación en diciembre. La modificación puntual del calendario laboral por el responsable del servicio exige 48 horas de antelación, máximo 5 jornadas al año, con compensación de 4,13 € por cambio."),
      art("27","Fiestas y descanso semanal: para personal fuera de turnos: sábados, domingos, festivos nacionales/autonómicos/locales y los días 24 y 31 de diciembre son inhábiles. El personal disfrutará durante la semana de San Isidro de 1 hora diaria de reducción o un día adicional de permiso. Se disfruta un «puente» anual. El personal de turnos que trabaje las noches del 24 o 31 de diciembre obtiene un día de permiso compensatorio."),
      art("28","Jornada continuada y partida: el personal fuera de turnos trabaja de lunes a viernes en jornada continuada o partida."),
      art("29","Periodos de descanso: 30 minutos en jornada continuada (computable como tiempo de trabajo efectivo). En jornada partida: descanso ininterrumpido de máximo 1,5 horas."),
      art("30","Horario, flexibilidad y control: fijado por la dirección informando al comité. Con jornada de 1.673,5 horas (actual): horario flexible 8:00-15:30 h con margen de ±30 min; permanencia obligatoria 8:30-14:30 h; recuperación de lunes a jueves entre 17:00-19:00 h. Con jornada de 1.562 horas (futura): horario 8:00-15:00 h con mismo margen; recuperación entre 16:30-18:30 h. El 5 de enero es recuperable. Flexibilidad ampliada hasta las 10:00 h para trabajadores con hijos hasta 14 años o ascendientes a partir de 70 años."),
      art("31","Servicios nocturnos: en los servicios nocturnos habrá siempre dos trabajadores de plantilla por razones de seguridad."),
      art("32","Retén: se establece para garantizar la seguridad e instalaciones ante siniestros o emergencias. El trabajador de retén debe estar localizable en la Comunidad de Madrid (o provincia colindante si allí tiene su domicilio) y presentarse en un máximo de 1 hora. El tiempo de disponibilidad no tiene la consideración de trabajo. Las intervenciones dan derecho a horas extraordinarias y a las indemnizaciones del art. 76.6."),
      art("33","Vacaciones: 25 días hábiles (fuera de turnos) o 32 días naturales (régimen de turnos); principalmente en junio-septiembre, hasta el 15 de enero del año siguiente; divisibles en máximo 7 periodos. La coincidencia con IT derivada de embarazo/parto/lactancia o con los permisos del art. 48 ET genera derecho a disfrutarlas en otra fecha. Prima de 290,84 € para turnos que disfruten vacaciones fuera de los períodos pico."),
      art("34","Horas extraordinarias: obligatorias cuando sean necesarias para garantizar el servicio esencial. Se compensan preferentemente con tiempo equivalente de descanso en los 4 meses siguientes. Las que no se compensen en tiempo se abonan al precio de la hora ordinaria, hasta un límite de 80 horas/año (a partir del cual todas se compensan en descanso). Las horas de fuerza mayor no computan en el máximo de 80."),

      art("35","Permisos y licencias retribuidos principales:"),
      bullet("20 días naturales por matrimonio o constitución de unión de hecho inscrita (empieza el día del hecho causante)."),
      bullet("1 hora diaria de lactancia hasta los 12 meses del menor (o reducción de media hora o acumulación en días completos)."),
      bullet("Exámenes prenatales y preparación al parto: tiempo imprescindible."),
      bullet("Hijos hospitalizados (hasta 18 años): reducción del 50% de jornada sin reducción salarial, máximo 15 días hábiles/año."),
      bullet("1 día natural por matrimonio de hijo."),
      bullet("Permiso parental: hasta 8 semanas (continuas o discontinuas) para el cuidado de menores hasta 8 años."),
      bullet("5 días naturales por fallecimiento de cónyuge, pareja de hecho, conviviente o parientes de 1.er grado (garantizados al menos 2 días laborables; 4 si requiere desplazamiento fuera de la provincia)."),
      bullet("2 días por fallecimiento de familiares de 2.º grado (4 si requiere desplazamiento)."),
      bullet("1 día natural por fallecimiento de familiares hasta 4.º grado (2 si requiere desplazamiento)."),
      bullet("5 días por accidente o enfermedad graves, hospitalización o intervención quirúrgica sin hospitalización con reposo de cónyuge/pareja/conviviente o parientes hasta 2.º grado."),
      bullet("Fuerza mayor familiar: 4 días retribuidos al año (horas retribuidas equivalentes)."),
      bullet("1 día natural por intervención quirúrgica sin hospitalización de padres, hijos, cónyuge, pareja o conviviente."),
      bullet("3 horas por consulta médica (ampliable si se acredita mayor duración)."),
      bullet("12 horas anuales para acompañar a hijos menores de 18 años al médico especialista (75 horas si hay alergias alimentarias severas)."),
      bullet("12 horas anuales para acompañar a padres a partir de 70 años o familiares dependientes hasta 2.º grado al médico especialista."),
      bullet("2 días por traslado de domicilio (disfrutables en los 3 meses siguientes)."),
      bullet("7 días de asuntos propios por año (periodo 1 julio – 30 junio; hasta 3 días fraccionables por horas)."),
      bullet("Tiempo imprescindible para exámenes académicos o profesionales y para deberes cívicos inexcusables."),
      bullet("Hasta 4 días por imposibilidad de acceder al centro de trabajo por catástrofe o fenómeno meteorológico adverso."),
      body("Suspensión del contrato (nacimiento/adopción/acogimiento): 19 semanas de suspensión (mejora respecto al ET) + 1 semana adicional de permiso retribuido junto al último periodo de suspensión."),
      body("Licencia sin sueldo: hasta 12 meses cada 2 años para el personal indefinido fijo que haya superado el periodo de prueba."),

      art("36","Excedencias: (1) Voluntaria: de 4 meses a 5 años (o hasta 10 si se acreditan 3 años de servicios). Requiere 15 días de preaviso. Derecho preferente al reingreso en vacante de igual o similar clasificación. (2) Forzosa: para desempeño de cargo público; reserva del puesto, cómputo de antigüedad y reingreso en el plazo de 2 meses desde el cese. (3) Especiales: hasta 3 años por cuidado de hijo (natural/adoptivo/acogido); hasta 2 años por cuidado de familiar incapacitado; hasta 1 año para prácticas curriculares; hasta 6 meses (prorrogables hasta 18) para víctimas de violencia de género. (4) Por servicios en el Grupo Canal: suspensión del contrato con reserva de reingreso al concluir la prestación en la empresa de destino."),
      sep(),

      // ═══════════════════════════
      // TÍTULO III — CLASIFICACIÓN PROFESIONAL
      // ═══════════════════════════
      h1("TÍTULO III — CLASIFICACIÓN PROFESIONAL"),
      art("37","Grupos profesionales: el personal se clasifica en 7 grupos y 3 áreas funcionales (Técnica, Administrativa, Operaria). Los grupos son: 1) Estructura y apoyo a estructura (subgrupos A y B); 2) Titulados universitarios (A y B); 3) Mandos intermedios (A y B); 4) Personal técnico (A, B, C y D); 5) Administrativos (A, B1 y B2); 6) Oficiales (A, B1 y B2); 7) Personal auxiliar (A y B)."),
      art("38","Factores de encuadramiento: conocimientos y experiencia, iniciativa, autonomía, responsabilidad, mando, complejidad y titulación."),
      art("39","Niveles académicos mínimos requeridos para el acceso exterior:"),
      bullet("Estructura y apoyo a estructura: titulación universitaria nivel 3 MECES (con formación/experiencia) o nivel 2 MECES con experiencia dilatada."),
      bullet("Titulados universitarios A: nivel 3 MECES; Titulados universitarios B: nivel 2 MECES."),
      bullet("Mandos intermedios, Personal técnico: FP grado superior (nivel 1 MECES)."),
      bullet("Administrativos A y B1: FP grado superior o bachillerato; B2: FP grado medio."),
      bullet("Oficiales A y B1: FP grado superior o medio; B2: FP grado medio, ESO o equivalente."),
      bullet("Personal auxiliar: FP grado medio, ESO, certificado de escolaridad o equivalente."),
      art("40","Definición de los grupos profesionales: se definen los 7 grupos con sus subgrupos en función del grado de autonomía, responsabilidad, mando, complejidad e iniciativa requeridos. Las funciones pueden realizarse en cualquier área de la empresa."),
      art("41","Puestos de trabajo: todos los puestos dispondrán de descripción. Sus actividades se recogerán en el Manual de Funciones (Disposición Adicional Cuarta)."),
      sep(),

      // ═══════════════════════════
      // TÍTULO IV — DESARROLLO PROFESIONAL, SELECCIÓN Y APRENDIZAJE
      // ═══════════════════════════
      h1("TÍTULO IV — DESARROLLO PROFESIONAL, SELECCIÓN Y APRENDIZAJE"),

      h2("Capítulo I — Desarrollo profesional y selección"),
      art("42","Principios generales: Canal aplica los principios constitucionales de igualdad, mérito y capacidad. La condición de indefinido fijo se adquiere superando los procesos selectivos. Orden de provisión de vacantes: 1.º turno de traslado; 2.º turno de promoción interna; 3.º turno de nuevo ingreso. Cupo de discapacidad: mínimo 7% de las vacantes de nuevo ingreso (al menos un 2% para discapacidad intelectual)."),
      art("43 a 44","Desarrollo profesional: se configura como elemento esencial de la política de RRHH. Modalidades: turno de traslado y turno de promoción interna."),
      art("45","Turno de traslado (concurso): para personal indefinido fijo con al menos 1 año en el puesto actual del mismo grupo, área funcional y análoga denominación. Puntuación: mejora de distancia (0,5 puntos/km, máx. 50 puntos), antigüedad en el destino (1,5 puntos/año) y antigüedad en el puesto (1,25 puntos/año, máx. 25 puntos). En caso de empate: 1.º discapacidad reconocida, 2.º trasladados forzosos, 3.º puntuación en prueba de conocimientos (si procede), 4.º antigüedad, 5.º mayor edad. La adjudicación es irrenunciable salvo fuerza mayor con acuerdo de la comisión."),
      art("46","Turno de promoción interna (concurso-oposición): para personal indefinido fijo en activo o en excedencia por cuidado de hijos/familiares/violencia de género. Pruebas: ejercicio tipo test (obligatorio, eliminatorio, máx. 10 puntos; superado con el 50%), prueba teórico-práctica (si procede, eliminatoria), evaluación de potencial y competencias (para titulados universitarios y mandos intermedios: prueba psicotécnica + entrevista personal, eliminatoria con 50%). En caso de empate: 1.º discapacidad, 2.º puntuación tipo test, 3.º prueba teórico-práctica, 4.º antigüedad, 5.º mayor edad."),
      art("47","Turno de nuevo ingreso (concurso-oposición): para candidaturas externas. Requisitos: no ser personal indefinido fijo de Canal, no haber sido despedido disciplinariamente, poseer la titulación y el carné B (cuando se exija). Pruebas idénticas a las de promoción interna con el añadido de una valoración de méritos (máx. 20% del total de pruebas). Lista de espera: vigencia 15 meses desde la publicación del proceso."),
      art("48 a 50","Selección temporal: se limita la contratación de duración determinada a lo estrictamente necesario. Las bolsas de empleo se forman con las listas de espera de procesos de nuevo ingreso, ordenadas por puntuación. El rechazo de un contrato temporal conlleva pasar al último puesto de la bolsa."),
      art("51 a 62","Disposiciones comunes: tramitación telemática; cupo de discapacidad del 7% (mínimo 2% intelectual); plazos de presentación de instancias (7 días hábiles para traslado/promoción, 12 días para nuevo ingreso); compensación económica a los miembros del tribunal: 60 €/prueba preparada, 70 €/sesión fuera de jornada. Reclamaciones en 2 días hábiles ante el tribunal calificador y, en segunda instancia, ante la dirección de RRHH."),

      h2("Capítulo II — Disposiciones comunes a las contrataciones"),
      art("63","Periodo de prueba: 6 meses para estructura/apoyo a estructura y titulados universitarios; 3 meses para el resto. Se interrumpe por IT, nacimiento, adopción, etc. No existe periodo de prueba si se superó anteriormente en Canal para el mismo grupo y especialidad."),
      art("64","Extinción/suspensión: preaviso de 20 días naturales para la dimisión voluntaria. Al causar baja el trabajador devuelve todos los elementos propiedad de la empresa."),

      h2("Capítulo III — Aprendizaje (Arts. 65-70)"),
      art("65","Principios: formación continua como factor clave del desarrollo profesional. Se aplica el modelo 70:20:10 (aprendizaje en el puesto, social y formal). Plan anual de formación con análisis previo de necesidades."),
      art("66","Modalidades formativas: corporativa, para el desempeño del puesto, para la promoción profesional (incluidas ayudas para estudios académicos), para la promoción personal y en materia de seguridad y salud laboral."),
      art("67","Comisión de formación paritaria (3+3) con reuniones trimestrales. Se promueve la paridad de género."),
      art("68","Asistencia y compensación: los cursos necesarios son obligatorios; la no asistencia injustificada puede ser sancionada. Los cursos fuera de jornada se compensan con horas equivalentes (excepto metodología online). Los gastos de desplazamiento y comida son asumidos por la empresa según normas internas."),
      art("69","Financiación: partida presupuestaria anual, incrementada con las subvenciones que concedan las administraciones competentes."),
      art("70","Formadores internos: complemento de 76 €/hora impartida fuera de jornada, 30 €/hora de preparación de material (máx. 50% de las horas del curso) y 30 €/hora de actualización de contenidos (máx. 33%). Requisito: mínimo 16 horas de formación/año y valoración media de los cursos superior a 3,5/5."),
      sep(),

      // ═══════════════════════════
      // TÍTULO V — RETRIBUCIONES ECONÓMICAS
      // ═══════════════════════════
      h1("TÍTULO V — RETRIBUCIONES ECONÓMICAS"),
      art("71","Principios generales: las retribuciones se fijan en base a jornada completa. El personal a tiempo parcial percibe las retribuciones en proporción a las horas trabajadas. El sistema retributivo del convenio sustituye y anula cualquier otro ingreso no contemplado en él, sin perjuicio del CGI."),
      art("72","Estructura salarial: a) Salario base; b) Complemento reclasificación; c) Complementos personales (antigüedad y complemento ad personam); d) Complementos de puesto de trabajo; e) Complementos de cantidad de trabajo (horas extraordinarias); f) Complementos de vencimiento superior al mes (pagas extra, complemento por objetivos, incentivo de productividad, complemento por desempeño y retribución media vacaciones). Las cuantías figuran en los Anexos I, II y III."),
      art("73","Salario base: corresponde al grupo y nivel profesional (Anexo I). Se abona en 12 mensualidades + 2 pagas extraordinarias."),
      art("74","Complemento reclasificación: corresponde al grupo y nivel profesional (Anexo II). Se abona en 12 mensualidades. Es revisable en el mismo porcentaje que el salario base."),
      art("75","Complementos personales:"),
      bullet("Antigüedad: 178,20 €/año de servicios prestados en la empresa. Revisable; no absorbible."),
      bullet("Complemento «ad personam»: cuantía resultante de la consolidación del complemento específico. Revisable al mismo porcentaje que el salario base; no absorbible salvo promoción que permita la absorción."),
      art("76","Complementos de puesto de trabajo:"),
      bullet("Plus de turnicidad — TURNO 1 (solo mañana o tarde o noche, o mañana+tarde): Titulados Universitarios 1.925,10 €/año; Técnicos y Mandos Intermedios 1.731,73 €/año; Administrativos y Oficiales 1.431,15 €/año; Auxiliar 1.302,36 €/año."),
      bullet("Plus de turnicidad — TURNO 2 (mañanas+tardes, tardes+noches o noches+mañanas, incl. festivos): Titulados 4.397,31 €; Técnicos y Mandos 3.941,37 €; Admtivos. y Oficiales 3.181,75 €; Auxiliar 2.801,94 €."),
      bullet("Plus de turnicidad — TURNO 3 (mañana, tarde y noche, incl. festivos): Titulados 4.438,41 €; Técnicos y Mandos 3.978,20 €; Admtivos. y Oficiales 3.211,49 €; Auxiliar 2.828,13 €."),
      bullet("Plus de noche (Turnos 2 y 3, por cada noche registrada en calendario): Titulados 29,38 €; Técnicos y Mandos 24,50 €; Admtivos. y Oficiales 19,57 €; Auxiliar 14,72 €."),
      bullet("Plus de trabajo en festivos: 8 €/sábado, domingo o festivo trabajado. Jornada nocturna en los días 24, 31 de diciembre y 5 de enero: 82 €."),
      bullet("Complemento específico: para estructura y apoyo a estructura según Anexo III (12 mensualidades). Se consolida a partir de 1 año de desempeño (30% el primer año + 10% por año sucesivo, hasta el 80%)."),
      bullet("Plus de disponibilidad (retén): 22,87 €/día laborable y 28,60 €/día no laborable. Si hay intervención efectiva: horas extraordinarias adicionales."),
      bullet("Plus de desplazamiento por retén: 76 € (vehículo propio desde localización hasta incidencia), 61 € (vehículo propio hasta el centro + vehículo empresa), 47,41 € (vehículo empresa desde localización)."),
      bullet("Complemento por atención comercial: Jefe de Equipo Admtvo. Comercial 3.246,96 €/año; Administrativo Comercial 1.608,12 €/año."),
      bullet("Complemento jornada partida: Titulados Sup. y Med. 1.665,36 €; Mandos y Técnicos 1.484,40 €; Oficiales, Admtivos. y Aux. Oficina/Especialistas 1.203,84 €; Peones 1.083,72 €."),
      bullet("Plus de secretaría: presidencia/DG/dirección 3.553,80 €/año; secretaría de subdirección 1.200 €/año."),
      bullet("Plus mandos intermedios: 804,06 €/año (excluido de la base de cálculo de variables anuales)."),
      bullet("Plus de caja: 804,06 €/año."),
      art("77","Horas extraordinarias: se abonan al precio de la hora ordinaria (calculado sobre salario base + complemento reclasificación + complementos personales + complementos fijos de puesto)."),
      art("78","Complementos de vencimiento superior al mes:"),
      bullet("Pagas extraordinarias: 2 al año (30 de junio y 22 de diciembre). Cuantía = 1 mensualidad de salario base. Devengo semestral proporcional."),
      bullet("Complemento por objetivos: para estructura, titulados A, técnicos A y mandos intermedios del área operaria: máximo 10% (12% para subdirecciones) sobre salario base + reclasificación + específico. Para el resto: máximo 1% en 2025, 1,25% en 2026, 1,75% en 2027 y 2,5% en 2028. Se abona en el primer trimestre del año siguiente."),
      bullet("Incentivo de productividad: 1.er año de contrato hasta 1,5%; 2.º año hasta 3,5%; a partir del 3.er año hasta 7% (base: salario base + reclasificación + específico). Se puede incrementar hasta 1,5% adicional si se lleva 5+ años en el puesto con valoración media de desempeño del 80%+; y otro 1,5% más si se lleva 9+ años con los mismos requisitos."),
      art("79","Indemnizaciones y suplidos: se incluyen los gastos de dietas, kilometraje y otras compensaciones específicas según normativa interna. El valor del km se actualizará periódicamente."),
      art("80","Cláusula de revisión salarial: incluye mecanismo de revisión anual conforme a los indicadores y condiciones pactadas en el convenio."),
      sep(),

      // ═══════════════════════════
      // TÍTULO VI — SEGURIDAD Y SALUD
      // ═══════════════════════════
      h1("TÍTULO VI — SEGURIDAD Y SALUD LABORAL"),
      art("81","Principios generales: las partes asumen el compromiso de garantizar la seguridad y salud conforme a la Ley 31/1995 de Prevención de Riesgos Laborales y sus normas de desarrollo."),
      art("82","Comité de seguridad y salud: composición y funciones según normativa vigente. Se dotará de los recursos necesarios para su funcionamiento efectivo."),
      art("83","Recursos preventivos: Canal designará los recursos preventivos exigidos por la normativa en los trabajos con especial peligrosidad."),
      art("84","Vigilancia de la salud: reconocimientos médicos periódicos. El personal adscrito a puestos que requieren reconocimiento obligatorio figura en el Anexo VII. Si del reconocimiento resulta una disminución de la capacidad laboral, el trabajador será reubicado en un puesto adecuado a su nueva capacidad (acoplamiento médico). Este procedimiento tiene prioridad en el orden de provisión de vacantes."),
      art("85","Fenómenos climatológicos adversos: se adoptarán medidas de protección y se aplicarán los protocolos de seguridad establecidos por las autoridades competentes."),
      art("86","Protección a la paternidad, maternidad y lactancia natural: las trabajadoras embarazadas, con parto reciente o en periodo de lactancia que ocupen puestos con riesgos específicos serán trasladadas a otro puesto sin riesgo. Si no es posible el traslado, se suspenderá el contrato con percepción de la prestación correspondiente."),
      art("87","Uniformidad y vestuario: la empresa proporcionará la ropa de trabajo y equipos de protección individual reglamentarios. El personal está obligado a su uso durante la jornada laboral."),
      art("88","Incapacidad temporal por accidente de trabajo y enfermedad profesional: se complementará hasta el 100% de las retribuciones desde el primer día."),
      art("89","Disminución de la capacidad laboral: cuando un trabajador vea reducida su capacidad para el desempeño de su puesto habitual, la empresa le asignará otro compatible con sus limitaciones, manteniendo sus retribuciones fijas."),
      sep(),

      // ═══════════════════════════
      // TÍTULO VII — PRESTACIONES ASISTENCIALES
      // ═══════════════════════════
      h1("TÍTULO VII — PRESTACIONES ASISTENCIALES Y COMPLEMENTARIAS"),

      h2("Capítulo I — Beneficios sociales"),
      art("90","Actividades deportivas y recreativas: la empresa facilitará el acceso a instalaciones deportivas y actividades de ocio para el personal y sus familias."),
      art("91","Abono de transporte: la empresa subvenciona el abono de transporte de la Comunidad de Madrid para el desplazamiento al trabajo. (Incompatibilidad con el plus de distancia/locomoción regulada en la Disposición Adicional Décima.)"),

      h2("Capítulo II — Acción complementaria"),
      art("92","Principio general: la empresa mejora la acción protectora de la Seguridad Social mediante las prestaciones complementarias del convenio."),
      art("93","Incapacidad temporal por enfermedad común o accidente no laboral: complemento hasta el 100% de las retribuciones desde el primer día, sujeto a los límites de absentismo previstos en la Disposición Adicional Quinta."),
      art("94","Nacimiento y cuidado de menor: además de la prestación legal, el convenio mejora la duración del permiso de nacimiento hasta 19 semanas (más 1 semana adicional de permiso retribuido junto al último periodo)."),

      h2("Capítulo III — Jubilación (Arts. 95-97)"),
      art("95","Extinción del contrato por jubilación: se estará a la normativa vigente de la Seguridad Social."),
      art("96","Jubilación parcial: se regulan las condiciones específicas para acceder a la jubilación parcial y la contratación del trabajador relevista."),
      art("97","Objetivos prioritarios de la política de empleo y jubilación obligatoria: se establece la posibilidad de jubilación obligatoria a los 67 años cuando se cumplan los requisitos legales y exista acuerdo en la comisión paritaria como medida vinculada a la política de empleo."),
      sep(),

      // ═══════════════════════════
      // TÍTULO VIII — DERECHOS DE REPRESENTACIÓN COLECTIVA
      // ═══════════════════════════
      h1("TÍTULO VIII — DERECHOS DE REPRESENTACIÓN COLECTIVA"),
      art("98","Competencias del comité de empresa: además de las legales, incluye: información sobre contrataciones temporales y subcontratación, consulta sobre medidas de conciliación, negociación de los calendarios laborales, participación en la comisión paritaria y en los tribunales de selección."),
      art("99","Derecho de reunión: el personal puede celebrar asambleas en los locales de la empresa fuera de la jornada laboral o, con autorización de la empresa, durante la jornada. Se regulan las condiciones de convocatoria y celebración."),
      art("100","Secciones sindicales: regulación de las secciones sindicales con representación en el comité de empresa, sus derechos, locales y crédito horario."),
      sep(),

      // ═══════════════════════════
      // TÍTULO IX — RÉGIMEN SANCIONADOR
      // ═══════════════════════════
      h1("TÍTULO IX — RÉGIMEN SANCIONADOR"),
      art("101","Principio general: el incumplimiento de las obligaciones laborales puede dar lugar a la imposición de sanciones, siempre mediante el oportuno expediente disciplinario."),

      art("102","Faltas — clasificación:"),
      h3("Faltas LEVES (prescripción: 10 días)"),
      bullet("Faltas de puntualidad: de 3 a 5 en un mes."),
      bullet("Falta de asistencia injustificada: 1 día/mes."),
      bullet("Negligencia o descuido en el trabajo."),
      bullet("No comunicar el traslado de domicilio."),
      bullet("Incorreción con compañeros, superiores o público."),
      bullet("No presentar el parte médico en el plazo de 4 días desde el inicio de la IT."),
      bullet("No asistir a cursos voluntarios."),
      bullet("No comunicar la carencia del permiso de conducir cuando sea exigible."),

      h3("Faltas GRAVES (prescripción: 20 días)"),
      bullet("Faltas de puntualidad: 20 o más en 3 meses."),
      bullet("Falta de asistencia injustificada: 2 o 3 días/mes."),
      bullet("Negligencia o desidia en el trabajo."),
      bullet("Desconsideración hacia superiores, compañeros o usuarios."),
      bullet("Indisciplina o desobediencia a las órdenes."),
      bullet("Imprudencia en el servicio."),
      bullet("Divulgación de secretos sin causar perjuicio."),
      bullet("Ocultar o no comunicar enfermedades que afecten al servicio."),
      bullet("No presentar el parte médico cuando la IT supere 4 días."),
      bullet("Incumplimiento de las obligaciones en materia de Seguridad Social."),
      bullet("No asistir a cursos necesarios para el puesto de trabajo."),
      bullet("Enfermedad grave que hubiera podido evitarse por negligencia."),

      h3("Faltas MUY GRAVES (prescripción: 60 días)"),
      bullet("Falta de asistencia injustificada: más de 3 días/mes."),
      bullet("20 o más faltas de puntualidad en 3 meses."),
      bullet("Fraude, deslealtad o abuso de confianza."),
      bullet("Causar daños graves a la empresa por dolo o negligencia grave."),
      bullet("Divulgación de secretos causando perjuicio."),
      bullet("Acoso (sexual, laboral, por razón de sexo, discriminatorio)."),
      bullet("Transgresión grave de la normativa de PRL."),
      bullet("Reincidencia en falta grave."),

      art("103","Sanciones:"),
      bullet("Faltas leves: amonestación verbal o escrita; suspensión de empleo y sueldo de 1 a 2 días."),
      bullet("Faltas graves: suspensión de empleo y sueldo de 3 a 15 días; suspensión del derecho a participar en pruebas selectivas durante 1 año."),
      bullet("Faltas muy graves: suspensión de empleo y sueldo de 16 días a 3 meses; traslado forzoso; despido disciplinario."),
      art("104","Procedimiento sancionador: notificación del pliego de cargos, audiencia al interesado y a los representantes de los trabajadores (para las muy graves), resolución motivada e impugnación ante la jurisdicción social."),
      art("105","Cancelación de las notas: leves a los 3 meses, graves a los 6 meses, muy graves al año; siempre que no se incurra en nuevas infracciones durante ese periodo."),
      art("106","Prescripción de las faltas: leves a los 10 días; graves a los 20 días; muy graves a los 60 días, contados desde que la empresa tuvo conocimiento de la falta y, en todo caso, a los 6 meses desde su comisión."),
      sep(),

      // ═══════════════════════════
      // DISPOSICIONES ADICIONALES (resumen)
      // ═══════════════════════════
      h1("DISPOSICIONES ADICIONALES Y TRANSITORIAS (resumen)"),
      h3("Disposiciones adicionales principales"),
      body("1.ª Compromiso de Garantías Individuales (CGI): garantía de mantenimiento de las condiciones económicas individuales superiores a las del convenio. Las diferencias se articulan como complemento personal CGI."),
      body("2.ª Reclasificación: regula la integración de los colectivos en la nueva estructura de clasificación del convenio y el complemento de reclasificación correspondiente."),
      body("3.ª Aplicación del CGI en cambios de puesto: el CGI se recalcula en los supuestos de cambio voluntario o involuntario de puesto, siguiendo las reglas de consolidación del complemento específico."),
      body("4.ª Manual de funciones y encuadramiento de puestos: la empresa elaborará y mantendrá actualizado el Manual de Funciones."),
      body("5.ª Minoración del complemento por desempeño en función del absentismo: el complemento por desempeño se ve reducido en función de los índices de absentismo individual, según los umbrales pactados."),
      body("6.ª Anticipos sobre el salario: regulación de las condiciones para la concesión de anticipos salariales."),
      body("7.ª Reconocimiento del complemento de antigüedad para personal sin CGI."),
      body("8.ª Antigüedad consolidada: consolidación de la antigüedad real para los efectos previstos en el convenio."),
      body("9.ª Permiso de conducción: obligación de mantener el carné B en vigor cuando sea requisito del puesto."),
      body("10.ª Incompatibilidad entre el plus de distancia/locomoción y la subvención del abono de transporte de la Comunidad de Madrid."),
      body("11.ª Modificaciones de horario: condiciones para la modificación puntual de horarios."),
      body("12.ª Teletrabajo: remisión al Anexo IX. Hasta un 30% de la jornada semanal en régimen de teletrabajo, de carácter voluntario y reversible, con dotación de medios por la empresa."),
      body("13.ª Desconexión digital: el personal tiene derecho a la desconexión digital fuera del horario de trabajo."),
      body("14.ª Bases de procesos selectivos: se crea una comisión ad hoc para debatir las características de los puestos que requieran pruebas específicas."),
      body("15.ª Tasas de reposición y procedimientos de traslado: aplicación de las tasas de reposición autorizadas por la legislación presupuestaria."),
      body("16.ª LGTBI: se aplican los principios y normas de la Ley 4/2023 y las leyes de la Comunidad de Madrid en materia de igualdad LGTBI."),
      body("17.ª Código ético y de conducta: remisión al código ético de Canal de Isabel II."),
      body("22.ª Efectos económicos: las mejoras económicas del convenio tendrán los efectos retroactivos desde el 1 de enero de 2025 según lo pactado."),
      body("23.ª Flexibilidad y cortesía: se establece el compromiso de atender con diligencia y cortesía al público en los servicios de atención al cliente."),
      body("24.ª Adaptaciones de jornada: posibilidad de adaptación de la duración y distribución de la jornada para facilitar la conciliación, en los términos del art. 34.8 ET."),

      h3("Disposiciones transitorias"),
      body("1.ª Reducción de los trabajos en turno de noche: planificación gradual para reducir la nocturnidad donde sea posible."),
      body("2.ª Evolución de los puestos de la clasificación profesional: revisión del encuadramiento de puestos en los nuevos grupos profesionales."),
      body("3.ª Productividad: régimen transitorio del complemento de productividad hasta la plena aplicación del nuevo sistema de incentivos del convenio."),
      body("4.ª Plus transitorio oficial corte de agua: mantenimiento del plus transitorio para los oficiales de corte de agua mientras no sea absorbido por el sistema retributivo general del convenio."),

      h3("Disposición final — Entrada en vigor"),
      body("El convenio entra en vigor el día 1 de enero de 2025, sin perjuicio de que determinadas materias tengan un momento de inicio de producción de efectos distinto, tal y como se indica en los artículos y disposiciones correspondientes."),
      sep(),

      // ═══════════════════════════
      // ANEXOS
      // ═══════════════════════════
      h1("ANEXOS (contenido resumido)"),
      body("ANEXO I — Tablas salariales: importe del salario base por grupo y nivel profesional.", true),
      body("ANEXO II — Complemento de reclasificación: importes del complemento de reclasificación por grupo y nivel.", true),
      body("ANEXO III — Complemento específico estructura: importes para los puestos de estructura y apoyo a estructura.", true),
      body("ANEXO IV — Grupos profesionales: listado descriptivo de los puestos de trabajo por grupo profesional y área funcional.", true),
      body("ANEXO V — Tabla de polivalencias: relación de puestos y las polivalencias funcionales aplicables entre ellos.", true),
      body("ANEXO VI — Compromiso de Garantías Individuales (CGI): texto del CGI, condiciones de su aplicación y forma de cálculo.", true),
      body("ANEXO VII — Puestos sujetos a reconocimiento médico obligatorio: listado de puestos que exigen vigilancia de la salud preceptiva.", true),
      body("ANEXO VIII — Protocolo de actuación frente al acoso psicológico en el trabajo: procedimiento de denuncia, instrucción y resolución de los expedientes de acoso.", true),
      body("ANEXO IX — Teletrabajo: condiciones de acceso, derechos y obligaciones del personal en régimen de trabajo a distancia. Voluntariedad, reversibilidad, dotación de medios por la empresa, derecho a la desconexión digital y régimen de prevención de riesgos.", true),
      esp(),
      new Paragraph({alignment:AlignmentType.CENTER,shading:{fill:AZUL3,type:ShadingType.CLEAR},
        children:[new TextRun({text:"Documento elaborado como resumen informativo. El texto oficial del III Convenio Colectivo de Canal de Isabel II, S.A.M.P. prevalece sobre este resumen a todos los efectos legales.",size:15,italic:true,font:"Arial",color:"555555"})],
        spacing:{before:300,after:100},indent:{left:300,right:300}})
    ]
  }]
});

Packer.toBuffer(doc).then(buf=>{
  fs.writeFileSync('resumen_III_convenio.docx', buf);
  console.log('Document created successfully.');
}).catch(e=>{
  console.error(e);
  process.exit(1);
});