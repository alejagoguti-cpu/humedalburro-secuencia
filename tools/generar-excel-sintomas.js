// Genera datos/red-sintomas-kennedy.xlsx (síntomas -> efectos de Kennedy) en el formato del visor.
// Uso: npm i exceljs && node tools/generar-excel-sintomas.js && node tools/build-red.js datos/red-sintomas-kennedy.xlsx
const ExcelJS = require('exceljs');
const path = require('path');
const ETIQ = require('./etiquetas-sintomas');
const OUT = path.join(__dirname, '..', 'datos', 'red-sintomas-kennedy.xlsx');

// ---- Lugares georreferenciados (lat, lon) -- OpenStreetMap/Nominatim y Wikipedia ----
const P = {
  CORA: [4.6302, -74.1588], VACA: [4.6287, -74.1620], VACA_SUR: [4.6239, -74.1658], BURRO: [4.6414, -74.1496],
  TECHO: [4.6452, -74.1417], TINTAL: [4.6439, -74.1538], PORTAL: [4.6294, -74.1724], PATIO: [4.6404, -74.1693],
  IGUALDAD: [4.6228, -74.1284], ALQ: [4.6014, -74.1384], HOSP: [4.6163, -74.1534]
};

const LAYERS = [
  ['C1', 'Economía territorial', 'Concentración de la oferta, informalidad y economía popular', '#FB923C', 'fa-coins', 'ECONOMÍA'],
  ['C2', 'Conocimiento e innovación', 'Brecha digital, residuos sin transformar e infraestructura social cerrada', '#A855F7', 'fa-lightbulb', 'INNOVACIÓN'],
  ['C3', 'Redes y flujos', 'Accesibilidad y robustez de vías, transporte y carga', '#38BDF8', 'fa-route', 'REDES'],
  ['C4', 'Metabolismo urbano', 'Agua, humedales, residuos, alimentos y suelo', '#4ADE80', 'fa-recycle', 'METABOLISMO'],
  ['C5', 'Efecto en el territorio', 'Consecuencias de cada síntoma: aparecen al abrir su sub-red', '#F43F5E', 'fa-bolt', 'EFECTO']
];
const TYPES = [
  ['T1', 'Efecto económico', 'El síntoma reduce ingresos, encarece costos o deja valor sin aprovechar.', '#FB923C', 'fa-coins'],
  ['T2', 'Efecto ambiental', 'El síntoma degrada agua, suelo, humedales o aire.', '#4ADE80', 'fa-leaf'],
  ['T3', 'Efecto en movilidad y redes', 'El síntoma retrasa, interrumpe o encarece flujos de personas y carga.', '#38BDF8', 'fa-route'],
  ['T4', 'Efecto social y de riesgo', 'El síntoma afecta ingresos, acceso a servicios o expone a la población a riesgo.', '#F43F5E', 'fa-people-group'],
  ['T5', 'Se refuerza por vía económica', 'Dos síntomas se alimentan entre sí a través de ingresos, empleo o costos.', '#FACC15', 'fa-link'],
  ['T6', 'Se refuerza por vía ambiental', 'Dos síntomas se agravan entre sí a través del agua, los residuos, el suelo o el aire.', '#A3E635', 'fa-water'],
  ['T7', 'Se refuerza por vía de movilidad y redes', 'Dos síntomas se agravan entre sí a través de las vías, el transporte o la carga.', '#2DD4BF', 'fa-bus'],
  ['T8', 'Se refuerza por vía social y de riesgo', 'Dos síntomas se agravan entre sí a través del acceso a servicios o de la exposición a riesgo.', '#E879F9', 'fa-heart-pulse'],
  ['T9', 'Comparten territorio (menos de 600 m)', 'Los dos síntomas ocurren a menos de 600 m de distancia: afectan a los mismos vecinos y calles. Relación derivada de las coordenadas.', '#94A3B8', 'fa-location-dot'],
  ['T10', 'Comparten entidad responsable', 'Una misma entidad (IDU, EAAB, UAESP, Secretaría de Movilidad, etc.) tiene competencia sobre los dos síntomas. Relación derivada de la columna Actores.', '#C4B5FD', 'fa-building-columns']
];

const Q_ECO = '¿El territorio genera oportunidades sin agotar recursos ni concentrar beneficios?';
const Q_INN = '¿El territorio puede adaptar su trayectoria con conocimiento e innovación?';
const Q_RED = '¿Los flujos y servicios son accesibles y robustos?';
const DOC = 'Documentado', INF = 'Inferido (verificar en campo)', GRP = 'Dato del grupo (verificar con fuente pública)';

// ---- SÍNTOMAS ----
// [id, nombre, capa, enfoque, lugar, lat/lon, qué se observa, por qué es síntoma, actores, pregunta, evidencia, fuente, sugerencia de imágenes]
const S = [
  ['S01', 'Corabastos concentra el abastecimiento de alimentos', 'Economía territorial', 'Economía · Redes', 'Central de Abastos, Av. Carrera 80 (UPZ Corabastos)', P.CORA,
    'Entre 5.000 y 11.000 toneladas de alimentos entran cada día a Corabastos según la fuente (la gerencia habla de 5.000 a 7.000 t; la prensa, de más de 10.000 t). La central ocupa 42 ha con 57 bodegas de venta y almacenamiento.',
    'Un solo punto concentra el flujo y la renta del abastecimiento de Bogotá y de Kennedy: lo que falla allí (bloqueo, paro, congestión) lo sufre toda la cadena.',
    'Corporación de Abastos de Bogotá, comerciantes mayoristas, transportadores, SDDE', Q_ECO, DOC,
    'La República 2022 (larepublica.co/economia/corabastos-es-la-nevera-de-bogota-movemos-aproximadamente-7-000-toneladas-al-dia-3404796); Wikipedia: Corabastos',
    'Foto 1: acceso principal por la Av. Carrera 80. Foto 2: bodegas de la central en hora pico.'],
  ['S02', 'Informalidad laboral alta en una localidad de 1,23 millones de personas', 'Economía territorial', 'Economía', 'Patio Bonito – borde de Corabastos', [4.6350, -74.1650],
    'Kennedy tiene 1.230.539 habitantes (2022) y 31.887 hab/km². La informalidad fuerte (sin prestaciones sociales) era de 48,0% entre las mujeres y 46,3% entre los hombres (Encuesta Multipropósito 2017); la Alcaldía la ubica entre las localidades con mayor informalidad (ECV 2023). El empleo informal y el rebusque se concentran alrededor de Corabastos.',
    'La economía local no genera suficiente empleo formal: gran parte del ingreso depende de oficios sin contrato ni seguridad social.',
    'Alcaldía Local de Kennedy, SDDE, comerciantes informales, hogares de Patio Bonito', Q_ECO, DOC,
    'Wikipedia: Kennedy (Bogotá); Alcaldía de Bogotá, ficha Kennedy destino de oportunidades (gobiernobogota.gov.co)',
    'Foto 1: comercio informal en el borde de Corabastos. Foto 2: calle de Patio Bonito en día de mercado.'],
  ['S03', '147 bodegas de reciclaje dispersas y mayoritariamente informales', 'Economía territorial', 'Economía · Metabolismo', 'Barrios alrededor de Corabastos (representado en el borde de la central)', [4.6255, -74.1565],
    'Kennedy tiene 147 bodegas de reciclaje (UAESP, 2012), la cifra más alta de Bogotá. Con tantas bodegas, el volumen por bodega es bajo, la operación cuesta más y baja el pago a los recicladores. La informalidad del reciclaje se estima en 79% (UAESP, citada en una tesis de 2023).',
    'El material que sostiene la economía circular de la localidad circula sin escala, sin registro y con márgenes mínimos.',
    'Recicladores de oficio, bodegueros, UAESP, ECA', Q_ECO, DOC,
    'Artículo con datos UAESP 2012 (revistas.lasalle.edu.co/index.php/ai/article/view/3369); tesis sobre La Alquería (repository.ugc.edu.co/handle/11396/8008)',
    'Foto 1: bodega de reciclaje típica. Foto 2: recicladores clasificando material.'],
  ['S04', 'Hogares vulnerables sin conexión estable a internet', 'Conocimiento e innovación', 'Innovación', 'Patio Bonito', P.PATIO,
    'El programa Conexión Social lleva fibra óptica de 25 Mbps a 16.000 hogares vulnerables de Kennedy. En Patio Bonito había familias que pagaban datos móviles para que los niños hicieran las tareas. La Alcaldía reconoce además la dificultad de la población vulnerable para acceder a empleos digitales de calidad.',
    'Sin conexión estable no hay acceso a formación, empleo digital ni emprendimiento: el conocimiento no llega a quienes más lo necesitan.',
    'Secretaría de Integración Social, ETB, hogares de Patio Bonito', Q_INN, DOC,
    'Publimetro, 27-ene-2026 (publimetro.co/noticias/2026/01/27/internet-gratis-en-los-hogares-de-bogota-la-estrategia-para-conectar-a-16000-familias-vulnerables); ImpactoTIC (impactotic.co/innovacion/innovacion-social/conexion-social-en-bogota-25-000-hogares-ya-tienen-internet-gratuito)',
    'Foto 1: hogar con instalación de fibra. Foto 2: estudiante haciendo tareas con datos móviles.'],
  ['S05', 'Residuos vegetales de Corabastos sin transformación a escala', 'Conocimiento e innovación', 'Innovación · Metabolismo', 'Central de Abastos', [4.6315, -74.1595],
    'Corabastos genera unas 51 t/día de residuos vegetales (estudio PNUD); otra fuente habla de 100 t/día de residuos y de que solo se aprovecha el 65% de la parte orgánica (UT Residuos Verdes). Una tesis de 2023 propone un biodigestor dentro de la central.',
    'El mayor generador de residuos orgánicos de la localidad aún no convierte ese flujo en compost, energía o insumos: se pierde valor y se carga el sistema de disposición.',
    'Corabastos S.A., UT Residuos Verdes, UAESP, Universidad EAN', Q_INN, DOC,
    'PNUD / SDA, gestión de residuos orgánicos en plazas de mercado (oab2.ambientebogota.gov.co); CONtexto Ganadero (contextoganadero.com); tesis EAN 2023',
    'Foto 1: contenedores de residuos vegetales. Foto 2: planta de transformación de residuos orgánicos.'],
  ['S06', 'Compostaje artesanal dentro del humedal La Vaca', 'Conocimiento e innovación', 'Innovación · Metabolismo', 'Humedal La Vaca (UPZ Corabastos)', [4.6287, -74.1625],
    'Los residuos vegetales del mantenimiento del humedal se compostan de manera muy artesanal, con proliferación de vectores y uso de químicos para controlar olores (tesis de la Universidad Antonio Nariño, 2021).',
    'Un ecosistema protegido resuelve sus residuos con técnicas improvisadas: falta método e innovación aplicada a la gestión del propio humedal.',
    'EAAB, SDA, comunidad del humedal, universidades', Q_INN, DOC,
    'Tesis UAN 2021 (repositorio.uan.edu.co/items/6123c6ac-f78f-45b3-a680-af5df2db2bc8)',
    'Foto 1: pila de compostaje en el humedal. Foto 2: borde del humedal con residuos vegetales.'],
  ['S07', 'Centro de reciclaje La Alquería cerrado', 'Conocimiento e innovación', 'Innovación · Economía', 'La Alquería La Fragua', P.ALQ,
    'El primer centro de reciclaje de Bogotá, ubicado en La Alquería (Kennedy), está cerrado indefinidamente y con la infraestructura deteriorada (tesis de 2023).',
    'La localidad con más bodegas de reciclaje no tiene operando su principal centro de aprovechamiento: la infraestructura de conocimiento y formalización está inactiva.',
    'UAESP, recicladores de oficio, ECA', Q_INN, DOC,
    'Tesis 2023 (repository.ugc.edu.co/handle/11396/8008); Las2orillas (las2orillas.co/?p=421645)',
    'Foto 1: fachada del centro cerrado. Foto 2: interior o patio deteriorado.'],
  ['S08', 'El rescate de alimentos es marginal frente al flujo de la central', 'Metabolismo urbano', 'Metabolismo · Economía', 'Centro de acopio del Banco de Alimentos en Corabastos', [4.6310, -74.1580],
    'El Banco de Alimentos rescata de 6 a 9 toneladas al día en Corabastos (676 comerciantes), frente a miles de toneladas que ingresan: menos del 0,2% del flujo. Bogotá pierde o desperdicia alrededor de 1,2 millones de t de alimentos al año (DNP).',
    'El sistema alimentario mueve enormes volúmenes pero recupera una fracción mínima: se agota el recurso sin aprovechar el excedente.',
    'Banco de Alimentos de Bogotá, comerciantes de Corabastos, SDDE', Q_ECO, DOC,
    'Banco de Alimentos de Bogotá (bancodealimentos.org.co); Alcaldía de Bogotá, estrategia contra la pérdida y el desperdicio de alimentos',
    'Foto 1: centro de acopio en Corabastos. Foto 2: entrega de alimentos rescatados.'],
  ['S09', 'Congestión crónica en la Av. Ciudad de Cali y la Av. Villavicencio', 'Redes y flujos', 'Redes', 'Portal Américas – Av. Ciudad de Cali con Av. Villavicencio', P.PORTAL,
    'En agosto de 2025 la congestión en la Av. Ciudad de Cali con Av. Villavicencio demoró la salida de buses del Portal Américas desde la madrugada. En junio de 2025, por obras en el corredor, TransMilenio reportó retrasos de hasta 30 minutos en la Troncal Américas.',
    'El corredor que conecta Kennedy con el resto de la ciudad no absorbe la demanda: el servicio no es robusto ante obras, accidentes o picos.',
    'IDU, TransMilenio, Secretaría de Movilidad, usuarios del Portal Américas', Q_RED, DOC,
    'Bogotá.gov.co, movilidad y TransMilenio 3-jun-2025 y 1-ago-2025 (bogota.gov.co/mi-ciudad/movilidad)',
    'Foto 1: fila de buses saliendo del Portal. Foto 2: congestión en la Av. Ciudad de Cali.'],
  ['S10', 'Calle 38 Sur deteriorada entre Corabastos y la Av. Ciudad de Cali', 'Redes y flujos', 'Redes', 'Calle 38 Sur entre Corabastos y la Av. Ciudad de Cali', [4.6285, -74.1650],
    'Un derecho de petición al IDU (enero de 2025) pide intervenir la Calle 38 Sur entre Corabastos y la Av. Ciudad de Cali; un vecino la describe como "una completa trocha".',
    'La vía local que conecta la central con la avenida principal está en mal estado: el tramo final de los flujos de carga y de las rutas zonales es el más frágil.',
    'IDU, vecinos, transportadores, Corabastos', Q_RED, DOC,
    'IDU, respuesta a derecho de petición, enero 2025 (idu.gov.co/Archivos_Portal/2025/servicio-a-la-ciudadania/consulte-sus-requerimientos/01-enero/03-01-25/202523600000011.pdf)',
    'Foto 1: calzada con huecos. Foto 2: camión transitando por el tramo.'],
  ['S11', 'Bloqueos en Patio Bonito interrumpen el transporte masivo', 'Redes y flujos', 'Redes', 'Av. Ciudad de Cali con Calle 38 Sur (Patio Bonito)', [4.6400, -74.1690],
    'En agosto de 2025 una manifestación en la Av. Ciudad de Cali con Calle 38 Sur obligó a desviar rutas zonales de TransMilenio. En septiembre de 2026 las manifestaciones en Patio Bonito llevaron a cerrar temporalmente el Portal Américas y varias estaciones.',
    'Un solo punto de la red, cuando se bloquea, deja sin servicio a varios barrios: la red de transporte tiene poca redundancia.',
    'TransMilenio, Secretaría de Movilidad, comunidad de Patio Bonito', Q_RED, DOC,
    'Bogotá.gov.co, TransMilenio 20-ago-2025 (publimetro.co); Publimetro, 1-sep-2026, manifestaciones en Patio Bonito',
    'Foto 1: bloqueo en la avenida. Foto 2: estación o portal cerrado.'],
  ['S12', 'Obras del Metro cierran por tramos el Portal Américas y la estación Biblioteca Tintal', 'Redes y flujos', 'Redes', 'Portal Américas y estación Biblioteca Tintal', [4.6300, -74.1732],
    'TransMilenio anunció cierres temporales del Portal Américas y de la estación Biblioteca Tintal por obras del Metro (noviembre de 2025; mayo y julio de 2026). La Línea 1 tiene 23,9 km y 75,5% de avance reportado en abril de 2026; su operación está prevista para marzo de 2028.',
    'Mientras llega el Metro, la capacidad del sistema actual baja: la transición depende de cierres que los usuarios de Kennedy absorben.',
    'Empresa Metro de Bogotá, TransMilenio, IDU', Q_RED, DOC,
    'Publimetro, 6-nov-2025 y 3-may-2026; Bogotá.gov.co, avance Línea 1 a feb-2026 (bogota.gov.co/mi-ciudad/movilidad/avance-de-la-linea-1-del-metro-de-bogota-con-corte-febrero-2026)',
    'Foto 1: obras del viaducto junto al Portal. Foto 2: plataforma cerrada.'],
  ['S13', 'Flujo de vehículos de carga hacia Corabastos sobre vías locales', 'Redes y flujos', 'Redes · Economía', 'Accesos a Corabastos (Av. Agoberto Mejía / Carrera 80)', [4.6322, -74.1565],
    'Un día de abril de 2020 ingresaron 641 vehículos con 4.211 toneladas (Alcaldía). La gerencia de Corabastos ha citado entre 14.000 y 15.000 vehículos diarios, cifra que no concuerda con las demás fuentes y debe verificarse.',
    'La carga de un mercado nacional entra por la malla vial de un barrio residencial: la red vial local sostiene un flujo que no fue diseñada para absorber.',
    'Corabastos, transportadores de carga, Secretaría de Movilidad', Q_RED, DOC + ' (cifra diaria sin consenso)',
    'Bogotá.gov.co, abastecimiento 2020 (bogota.gov.co/mi-ciudad/desarrollo-economico/este-lunes-ingresaron-bogota-4211-toneladas-de-alimentos); La República 2022',
    'Foto 1: camiones haciendo fila en el acceso. Foto 2: zona de cargue y descargue.'],
  ['S14', 'El humedal El Burro está partido por la Av. Ciudad de Cali', 'Metabolismo urbano', 'Metabolismo', 'Humedal El Burro', P.BURRO,
    'El Burro conserva unas 18,8 ha tras perder 89% de su superficie en 70 años. La Av. Ciudad de Cali lo atraviesa y lo divide en dos sectores (80% oriental y 20% occidental, según un artículo), sin conexión hídrica entre ellos. Tiene 33 especies de aves registradas.',
    'Una vía de movilidad se impuso sobre un sistema hídrico: el humedal pierde su función de regulación y conectividad.',
    'EAAB, SDA, IDU, colectivos ambientales', Q_RED, DOC,
    'Wikipedia: El Burro; El Espectador (elespectador.com/bogota/el-humedal-el-burro-perdio-el-89-de-su-ecosistema-por-la-urbanizacion-de-bogota-article-891072)',
    'Foto 1: vista aérea del humedal con la avenida. Foto 2: sector occidental.'],
  ['S15', 'El humedal La Vaca ocupado y con residuos y conexiones erradas', 'Metabolismo urbano', 'Metabolismo', 'Humedal La Vaca (entre El Amparo y Corabastos)', [4.6270, -74.1610],
    'Se reporta cerca del 95% de su suelo ocupado por invasiones, que comenzaron en los años 70. Botellas, bolsas, electrodomésticos y muebles se arrojan al humedal y a los sumideros; hay conexiones erradas de alcantarillado. Tiene entre 7 y 8 ha.',
    'El humedal vecino de la central mayorista funciona como borde de descarga: pierde superficie y capacidad de limpiar el agua.',
    'EAAB, SDA, comunidad, Corabastos', Q_RED, DOC,
    'Wikipedia: Humedal La Vaca',
    'Foto 1: borde del humedal con vivienda. Foto 2: residuos en el cuerpo de agua.'],
  ['S16', 'El humedal de Techo fragmentado por vías y ocupado', 'Metabolismo urbano', 'Metabolismo', 'Humedal de Techo', P.TECHO,
    'Mide entre 11 y 11,67 ha. La Carrera 80 y la Av. Agoberto Mejía lo dividen en un sector norte y otro sur; el barrio Lagos de Castilla lo ocupó desde los años 90 y el sector norte alberga industrias y parqueaderos.',
    'Las vías y la ocupación fragmentan el humedal: el sistema natural que debería amortiguar el agua de la zona queda en pedazos.',
    'EAAB, SDA, IDU, Lagos de Castilla', Q_RED, DOC,
    'Wikipedia: Humedal de Techo',
    'Foto 1: carrera 80 junto al humedal. Foto 2: sector norte con parqueaderos.'],
  ['S17', 'Conexiones erradas y desbordes en el río Fucha', 'Metabolismo urbano', 'Metabolismo', 'Barrio La Igualdad, río Fucha', P.IGUALDAD,
    'La EAAB anunció la eliminación de 285 conexiones erradas en Kennedy (inversión citada de 8.000 millones de pesos; la fecha de la nota debe verificarse). El río Fucha a la altura del barrio La Igualdad es uno de los 107 puntos críticos de Bogotá por riesgo de inundación por desbordamiento.',
    'Las aguas residuales llegan por las redes de aguas lluvias a los cuerpos de agua y el río no tiene capacidad de absorber crecientes: el sistema no es robusto.',
    'EAAB, IDIGER, Secretaría de Ambiente', Q_RED, DOC,
    'Bogotá.gov.co, Acueducto cierra conexiones erradas (bogota.gov.co/mi-ciudad/habitat/acueducto-cierra-conexiones-de-agua-que-contaminaban-rios-de-bogota); operativos de inundación (bogota.gov.co/mi-ciudad/ambiente/operativos-en-bogota-para-prevenir-inundaciones-en-fonomeno-de-la-nina)',
    'Foto 1: río Fucha a su paso por La Igualdad. Foto 2: tubería de conexión errada.'],
  ['S18', 'El 35% de Kennedy es terreno inundable', 'Metabolismo urbano', 'Metabolismo', 'Valle aluvial del río Bogotá, occidente de Kennedy (representado en el sector de La Vaca Sur)', [4.6185, -74.1700],
    'Un documento de la Alcaldía Local indica que Kennedy es plana, con pequeñas depresiones, y que el 35% de su área es inundable por estar por debajo de las posibilidades de desagüe natural. Pertenece al valle aluvial del río Bogotá, que recibe al Fucha y al Tunjuelo.',
    'Una localidad de 1,2 millones de habitantes está asentada sobre suelo que no drena: el riesgo es estructural, no ocasional.',
    'Alcaldía Local de Kennedy, IDIGER, EAAB', Q_RED, DOC,
    'Alcaldía Local de Kennedy, acta/documento (redjurista.com, a_jalkennedy_0003_2021)',
    'Foto 1: calle inundada en temporada de lluvias. Foto 2: mapa de zonas inundables.'],
  ['S19', 'Déficit de espacio público y zonas verdes', 'Metabolismo urbano', 'Metabolismo', 'Barrios consolidados del centro de Kennedy (representado en el entorno del Hospital de Kennedy)', [4.6200, -74.1500],
    'Greenpeace (2020) incluye a Kennedy entre las 13 de 19 localidades con déficit de áreas verdes (4 a 8 m² por habitante frente a 10 m² de referencia). Bogotá tiene 3,9 m² de espacio público efectivo por habitante frente a 15 m² de la norma; la SDP señala déficit importante de plazas en Kennedy.',
    'Con 31.887 hab/km² faltan lugares públicos para el comercio local, el encuentro y la sombra: el espacio libre es el servicio más escaso.',
    'DADEP, IDRD, SDP, Alcaldía Local', Q_RED, DOC,
    'Greenpeace Colombia, déficit de áreas verdes (greenpeace.org/colombia); SDP / Caja de Vivienda Popular (sistemas002.sdp.gov.co)',
    'Foto 1: parque con poca vegetación. Foto 2: andén ocupado por comercio.'],
  ['S20', 'Carga orgánica y lixiviados en el borde de Corabastos', 'Metabolismo urbano', 'Metabolismo', 'Canales de vertimiento contiguos a Corabastos', [4.6298, -74.1600],
    'La memoria técnica del grupo registra una DBO₅ superior a 180 mg/L en los canales de vertimiento junto a Corabastos, además de ruido de 75 a 84 dB(A) en la Av. Ciudad de Cali y 78,4% de suelo sellado (NDBI). Estos valores provienen del trabajo del grupo y deben contrastarse con mediciones oficiales.',
    'El metabolismo de la central se descarga sobre el borde del humedal: el agua pierde oxígeno y la fauna pierde hábitat.',
    'Corabastos, EAAB, SDA', Q_RED, GRP,
    'Memoria técnica Corte II del grupo (memoria-completa-chat-kennedy-corte2.md)',
    'Foto 1: canal junto a Corabastos. Foto 2: punto de muestreo.']
];

// ---- EFECTOS ----
// [id, nombre, enfoque, lugar, lat/lon, qué ocurre, por qué importa, evidencia, fuente, sugerencia de imágenes]
const E = [
  ['E01', 'Riesgo de desabastecimiento por depender de una sola central', 'Economía · Redes', 'Corabastos', [4.6300, -74.1595],
    'Cuando se bloquean los accesos o hay paro camionero, la oferta de la central baja y los precios suben: lo reportaron los comerciantes de Corabastos durante el paro camionero.',
    'Pasa de síntoma a pérdida real: la red de abastecimiento no es robusta porque no tiene un segundo punto de entrada.', DOC, 'El Tiempo, paro camionero y Corabastos (eltiempo.com/bogota/precios-por-las-nubes-y-escasez-de-alimentos-asi-viven-los-comerciantes-de-corabastos-las-afectaciones-del-paro-camionero-en-bogota-3378480)', 'Foto 1: puestos con poca oferta. Foto 2: camiones detenidos.'],
  ['E02', 'Tráfico pesado sobre vías de borde de la central', 'Redes', 'Av. Agoberto Mejía / Calle 38 Sur', [4.6330, -74.1555],
    'Los vehículos de carga comparten la Av. Agoberto Mejía y la Calle 38 Sur con el tránsito local, sobre calzadas deterioradas.',
    'Cada camión retrasado en el borde de la central retrasa también a los buses zonales y al peatón: el flujo de carga y el de personas compiten por la misma vía.', INF, 'IDU 2025 (petición Calle 38 Sur); Alcaldía 2020', 'Foto 1: camión en vía local. Foto 2: cruce con peatones.'],
  ['E03', 'Ingresos inestables y sin seguridad social', 'Economía', 'Patio Bonito', [4.6398, -74.1688],
    'Quien trabaja en informalidad no cotiza a salud ni pensión y su ingreso depende de cada día de trabajo.',
    'La economía local crece en población pero no en protección: un mes malo afecta al hogar completo.', DOC, 'Alcaldía de Bogotá, ficha Kennedy 2026', 'Foto 1: vendedor ambulante. Foto 2: hogar de Patio Bonito.'],
  ['E04', 'Ingresos del borde de la central dependen de que la central no se detenga', 'Economía', 'Borde de Corabastos', [4.6335, -74.1612],
    'Cargueros, vendedores y recicladores del entorno viven de la actividad diaria de la central: si hay paro, bloqueo o cierre, el ingreso cae ese mismo día.',
    'La economía popular del sector no tiene colchón: depende de un único nodo que no controla.', INF, 'Wikipedia: Kennedy (Bogotá); Alcaldía 2026', 'Foto 1: cargueros en el acceso. Foto 2: vendedores en el borde.'],
  ['E05', 'Recicladores con ingresos bajos por volumen fragmentado', 'Economía', 'Barrios alrededor de Corabastos', [4.6262, -74.1555],
    'Con 147 bodegas, cada una recupera poco material; eso encarece la operación y reduce el pago a los recicladores.',
    'La riqueza del residuo se pierde en la dispersión: se agota el recurso sin que el ingreso llegue a quien lo recupera.', DOC, 'Artículo con datos UAESP 2012 (revistas.lasalle.edu.co/index.php/ai/article/view/3369)', 'Foto 1: reciclador con carreta. Foto 2: material clasificado.'],
  ['E06', 'Material reciclado sin trazabilidad ni venta formal', 'Economía · Innovación', 'Kennedy', [4.6272, -74.1540],
    'Con 79% de informalidad estimada, el material recuperado no se registra ni se vende como cadena formal.',
    'Sin registro no se pueden medir toneladas, ni acceder a incentivos ni financiar mejoras: la innovación en reciclaje no tiene datos con qué empezar.', INF, 'UAESP citada en tesis 2023', 'Foto 1: pesaje informal. Foto 2: bodega sin registro.'],
  ['E07', 'Estudiantes sin conexión estable para hacer tareas', 'Innovación', 'Patio Bonito', [4.6392, -74.1680],
    'Las familias pagaban datos móviles para que los niños hicieran las tareas escolares.',
    'El costo de conectarse compite con la comida y el transporte: el aprendizaje queda sujeto al saldo del celular.', DOC, 'Publimetro, 27-ene-2026', 'Foto 1: niño con celular. Foto 2: tarea en la mesa.'],
  ['E08', 'Menos acceso a empleo digital y cursos de emprendimiento', 'Innovación · Economía', 'Patio Bonito', [4.6418, -74.1702],
    'La población vulnerable tiene dificultad para acceder a empleos digitales de calidad; el programa Conexión Social agrega una plataforma con cursos de emprendimiento y trabajo.',
    'Quien no tiene conexión queda fuera de los cursos y de las ofertas: la brecha digital se vuelve brecha de ingresos.', DOC, 'Alcaldía de Bogotá, certificación laboral; ImpactoTIC 2026', 'Foto 1: aula de formación. Foto 2: persona en curso en línea.'],
  ['E09', 'El 35% de la fracción orgánica de Corabastos no se aprovecha', 'Metabolismo · Innovación', 'Corabastos', [4.6312, -74.1603],
    'De la parte orgánica de los residuos de la central solo se usa el 65%; el resto no se aprovecha.',
    'Cada tonelada sin tratar es transporte, disposición y emisiones que podrían evitarse con compostaje o biodigestión.', DOC, 'CONtexto Ganadero (UT Residuos Verdes)', 'Foto 1: residuos acumulados. Foto 2: camión recolector.'],
  ['E10', 'Compost y biogás que no se producen', 'Economía · Innovación', 'Corabastos', [4.6295, -74.1580],
    'Una tesis de la Universidad EAN (2023) propone un biodigestor dentro de la central para generar biogás; sin esa infraestructura, el valor energético y agrícola del residuo sale como desecho.',
    'Oportunidad económica perdida: el residuo de la central podría abastecer insumos o energía a la propia localidad.', DOC, 'Universidad EAN 2023 (repository.universidadean.edu.co)', 'Foto 1: biodigestor de referencia. Foto 2: montones de residuo.'],
  ['E11', 'Vectores y olores en el humedal La Vaca', 'Metabolismo', 'Humedal La Vaca', [4.6280, -74.1632],
    'El compostaje artesanal atrae vectores y produce olores en un humedal rodeado de barrios.',
    'Un problema de gestión de residuos se convierte en problema de salud ambiental para quienes viven al lado.', DOC, 'Tesis UAN 2021', 'Foto 1: pila de material vegetal. Foto 2: vecino junto al humedal.'],
  ['E12', 'Uso de químicos para controlar olores dentro de un humedal', 'Metabolismo', 'Humedal La Vaca', [4.6296, -74.1645],
    'Para controlar los olores del compostaje se usan químicos en un ecosistema que debería depurar el agua.',
    'El remedio agrega contaminación al sistema que se quiere recuperar.', DOC, 'Tesis UAN 2021', 'Foto 1: aplicación de químicos. Foto 2: agua del humedal.'],
  ['E13', 'Recicladores sin una sede formal en la localidad con más bodegas', 'Economía · Innovación', 'La Alquería', [4.6021, -74.1397],
    'Con el centro de La Alquería cerrado, los recicladores de Kennedy no cuentan con ese punto formal para acopiar y vender.',
    'Se pierde el lugar donde la formalización y la capacitación podían ocurrir.', INF, 'Tesis 2023; Las2orillas', 'Foto 1: reciclador en la calle. Foto 2: fachada cerrada.'],
  ['E14', 'Infraestructura pública deteriorada por abandono', 'Economía', 'La Alquería', [4.6008, -74.1378],
    'El centro, que fue el primero de la ciudad, está cerrado indefinidamente y con infraestructura deteriorada.',
    'Una inversión pública pierde valor cada año que permanece sin uso.', DOC, 'Tesis 2023 (repository.ugc.edu.co/handle/11396/8008)', 'Foto 1: fachada. Foto 2: patio.'],
  ['E15', 'Se pierde comida mientras hay hogares con inseguridad alimentaria', 'Economía · Metabolismo', 'Corabastos', [4.6307, -74.1572],
    'Bogotá pierde o desperdicia alrededor de 1,2 millones de t de alimentos al año; en el país casi un tercio de los hogares vive en inseguridad alimentaria grave o moderada.',
    'El excedente existe pero no llega a quien lo necesita: el flujo de alimentos es abundante y a la vez poco accesible.', DOC, 'Alcaldía de Bogotá (DNP); EFE, oct-2023', 'Foto 1: alimentos descartados. Foto 2: entrega a comedor comunitario.'],
  ['E16', 'Retrasos de hasta 30 minutos en la Troncal Américas', 'Redes', 'Portal Américas', [4.6290, -74.1716],
    'Por obras en el corredor, TransMilenio informó retrasos de hasta 30 minutos en el servicio.',
    'Quienes viven en Kennedy gastan tiempo y dinero extra en cada viaje al trabajo: la red de transporte no es robusta.', DOC, 'Bogotá.gov.co, 3-jun-2025', 'Foto 1: estación llena. Foto 2: tablero de servicio con demora.'],
  ['E17', 'Accidentes y siniestros viales sobre la avenida', 'Redes', 'Av. Ciudad de Cali con Calle 6D (Biblioteca El Tintal)', [4.6439, -74.1545],
    'En agosto de 2025 un siniestro vial en la Av. Ciudad de Cali con Calle 6D, junto a la Biblioteca El Tintal, generó gran congestión.',
    'Un incidente aislado colapsa un tramo entero: no hay vías alternas que absorban el flujo.', DOC, 'Bogotá.gov.co, movilidad 1-ago-2025', 'Foto 1: siniestro en la avenida. Foto 2: congestión posterior.'],
  ['E18', 'Calzadas deterioradas encarecen y frenan la carga y las rutas zonales', 'Redes · Economía', 'Calle 38 Sur', [4.6278, -74.1642],
    'La Calle 38 Sur está descrita como "una completa trocha" entre Corabastos y la Av. Ciudad de Cali.',
    'El tramo final de la cadena de abastecimiento es el más lento y el que más desgasta vehículos.', DOC, 'IDU, respuesta a petición, enero 2025', 'Foto 1: huecos en la vía. Foto 2: vehículo atascado.'],
  ['E19', 'Rutas zonales desviadas', 'Redes', 'Patio Bonito', [4.6425, -74.1690],
    'Una manifestación en la Av. Ciudad de Cali con Calle 38 Sur obligó a desviar rutas zonales.',
    'Los barrios dependen de una sola avenida: cuando se cierra, quedan sin servicio.', DOC, 'Bogotá.gov.co, 20-ago-2025', 'Foto 1: bus desviado. Foto 2: usuarios esperando.'],
  ['E20', 'Portal Américas cerrado temporalmente', 'Redes', 'Portal Américas', [4.6298, -74.1730],
    'Las manifestaciones en Patio Bonito y las obras del Metro han llevado a cerrar temporalmente el portal y varias estaciones.',
    'El principal acceso al transporte masivo del occidente de Kennedy queda inutilizado por periodos: es un punto único de falla.', DOC, 'Publimetro, 1-sep-2026 y 6-nov-2025', 'Foto 1: portal con cierre. Foto 2: usuarios buscando alternativa.'],
  ['E21', 'Capacidad del sistema reducida durante las obras del Metro', 'Redes', 'Estación Biblioteca Tintal', [4.6436, -74.1548],
    'La estación Biblioteca Tintal y el Portal Américas cierran por tramos mientras avanza la Línea 1 (75,5% a abril de 2026; operación prevista en 2028).',
    'La transición hacia una mejor red cuesta años de servicio reducido para quienes más lo usan.', DOC, 'Publimetro, 6-nov-2025; Bogotá.gov.co, avance Línea 1', 'Foto 1: obra del viaducto. Foto 2: estación cerrada.'],
  ['E22', 'Aves y agua no pasan entre los dos sectores de El Burro', 'Metabolismo', 'Humedal El Burro', [4.6420, -74.1508],
    'La avenida corta la conectividad hídrica y el paso de aves entre el sector oriental (80%) y el occidental (20%).',
    'Se pierde la función de regulación del humedal y su capacidad de recuperarse.', DOC, 'Artículo sobre el humedal El Burro; Wikipedia: El Burro', 'Foto 1: avenida sobre el humedal. Foto 2: ave en el sector oriental.'],
  ['E23', 'Humedal reducido a una fracción: 89% de su superficie perdida', 'Metabolismo', 'Humedal El Burro', [4.6408, -74.1488],
    'En 70 años El Burro perdió 89% de su superficie; hoy tiene unas 18,8 ha y 33 especies de aves registradas.',
    'Menos espacio para absorber lluvias y para el hábitat: el barrio queda más expuesto a inundaciones.', DOC, 'El Espectador; Universidad Nacional citada por El Tiempo', 'Foto 1: foto histórica del humedal. Foto 2: borde urbano actual.'],
  ['E24', 'La Vaca pierde superficie y capacidad de amortiguar crecidas', 'Metabolismo', 'Humedal La Vaca', [4.6275, -74.1618],
    'Con cerca del 95% del suelo ocupado, el humedal tiene poca área libre para almacenar agua.',
    'Una crecida encuentra menos espacio y se desplaza a calles y viviendas del entorno.', INF, 'Wikipedia: Humedal La Vaca', 'Foto 1: borde con vivienda. Foto 2: sector sur.'],
  ['E25', 'Residuos y aguas negras entran al drenaje y a los humedales', 'Metabolismo', 'Humedales La Vaca y Techo', [4.6242, -74.1662],
    'Botellas, bolsas y muebles llegan a los sumideros; las conexiones erradas descargan agua residual en el drenaje de lluvias.',
    'El sistema que debía evacuar la lluvia se tapa y se contamina: sube el riesgo de encharcamiento.', DOC, 'Wikipedia: Humedal La Vaca; EAAB', 'Foto 1: sumidero tapado. Foto 2: descarga en el humedal.'],
  ['E26', 'El humedal de Techo podría quedar en unas 3 ha', 'Metabolismo', 'Humedal de Techo', [4.6458, -74.1408],
    'Según una fuente citada, su futuro podría limitarse a unas tres hectáreas por el impacto de las construcciones ilegales.',
    'La pérdida sería irreversible: un servicio de regulación que no se puede reconstruir.', DOC + ' (proyección citada)', 'Wikipedia: Humedal de Techo', 'Foto 1: sector norte. Foto 2: mapa de ocupación.'],
  ['E27', 'Aguas contaminadas llegan a los ríos Fucha y Bogotá', 'Metabolismo', 'Cuenca del Fucha', [4.6240, -74.1300],
    'Una conexión errada lleva aguas residuales a las tuberías de aguas lluvias, que descargan en ríos y humedales.',
    'La contaminación de la central y de los barrios sale del territorio y se acumula aguas abajo.', DOC, 'EAAB (conexiones erradas)', 'Foto 1: descarga al río. Foto 2: muestra de agua.'],
  ['E28', 'Desbordamiento del Fucha en La Igualdad', 'Metabolismo', 'Barrio La Igualdad', [4.6230, -74.1288],
    'El tramo del río Fucha en La Igualdad es uno de los 107 puntos críticos de Bogotá por riesgo de desbordamiento.',
    'Las viviendas junto al río quedan expuestas a pérdidas cuando llueve fuerte.', DOC, 'Bogotá.gov.co, operativos de inundación', 'Foto 1: río en crecida. Foto 2: viviendas junto a la ronda.'],
  ['E29', 'Inundaciones históricas en Patio Bonito', 'Metabolismo · Economía', 'Patio Bonito', [4.6395, -74.1700],
    'El sector sufrió inundaciones graves en 1979 y 1980; la zona era parte del cauce del río Bogotá.',
    'El barrio que nació junto a la central se construyó sobre suelo inundable: el riesgo se hereda.', DOC, 'Wikipedia: Patio Bonito', 'Foto 1: foto histórica de la inundación. Foto 2: barrio actual.'],
  ['E30', 'Calor superficial de hasta 33 °C con 78% de suelo sellado', 'Metabolismo', 'Centro de Kennedy', [4.6205, -74.1490],
    'La memoria del grupo registra 78,4% de suelo cubierto por asfalto y concreto y temperaturas superficiales de 28,5 a 33,2 °C.',
    'Menos sombra y más calor en barrios que ya tienen pocas zonas verdes (dato del grupo; contrastar con fuente oficial).', GRP, 'Memoria técnica Corte II del grupo', 'Foto 1: imagen térmica o mapa. Foto 2: calle sin arbolado.'],
  ['E31', 'Menor oxígeno en el agua del borde de la central', 'Metabolismo', 'Humedal La Vaca, borde de Corabastos', [4.6300, -74.1615],
    'Con DBO₅ superior a 180 mg/L en los canales de vertimiento, el agua pierde oxígeno y favorece bacterias como E. coli (dato del grupo).',
    'Menos oxígeno significa menos vida acuática y un humedal menos capaz de limpiar el agua.', GRP, 'Memoria técnica Corte II del grupo', 'Foto 1: canal junto a la central. Foto 2: toma de muestra.']
];

// ---- ARISTAS: [síntoma/origen, destino, tipo, tensión, peso, justificación] ----
const T = { 1: 'Efecto económico', 2: 'Efecto ambiental', 3: 'Efecto en movilidad y redes', 4: 'Efecto social y de riesgo', 5: 'Se refuerza por vía económica', 6: 'Se refuerza por vía ambiental', 7: 'Se refuerza por vía de movilidad y redes', 8: 'Se refuerza por vía social y de riesgo', 9: 'Comparten territorio (menos de 600 m)', 10: 'Comparten entidad responsable' };
const W = { 'Crítica': 4, 'Severa': 3, 'Alta': 2, 'Media': 1 };
const ED = [
  ['S01', 'E01', 1, 'Severa', 'Como casi todo el abastecimiento pasa por un punto, cualquier bloqueo o paro lo reduce y sube los precios.'],
  ['S01', 'E02', 3, 'Alta', 'Miles de toneladas diarias se mueven en vehículos que entran por las vías de borde de la central.'],
  ['S01', 'E15', 4, 'Alta', 'Un flujo tan grande genera excedentes y pérdidas que no llegan a los hogares que los necesitan.'],
  ['S02', 'E03', 4, 'Alta', 'El empleo informal no incluye cotización a salud ni pensión.'],
  ['S02', 'E04', 1, 'Alta', 'Muchos ingresos informales dependen de la actividad diaria alrededor de la central.'],
  ['S02', 'E08', 1, 'Media', 'Sin empleo formal la población tiene menos acceso a formación y a ofertas de trabajo calificado.'],
  ['S03', 'E05', 1, 'Alta', 'Con 147 bodegas el volumen por bodega es bajo y baja el pago a los recicladores.'],
  ['S03', 'E06', 1, 'Media', 'La informalidad impide registrar el material y venderlo como cadena formal.'],
  ['S03', 'E13', 1, 'Media', 'Sin escala formal los recicladores no cuentan con una sede propia de acopio.'],
  ['S04', 'E07', 4, 'Alta', 'Las familias pagan datos móviles para tareas escolares porque no tienen conexión fija.'],
  ['S04', 'E08', 1, 'Alta', 'La brecha de conexión limita el acceso a cursos de emprendimiento y a empleo digital.'],
  ['S05', 'E09', 2, 'Severa', 'Solo el 65% de la fracción orgánica se aprovecha; el resto queda sin tratamiento.'],
  ['S05', 'E10', 1, 'Alta', 'Sin biodigestor ni planta de compostaje, el valor del residuo no se convierte en insumo ni energía.'],
  ['S06', 'E11', 2, 'Alta', 'El compostaje artesanal atrae vectores y olores en un humedal rodeado de barrios.'],
  ['S06', 'E12', 2, 'Media', 'Para controlar el olor se aplican químicos dentro del propio humedal.'],
  ['S07', 'E13', 1, 'Media', 'Con el centro cerrado, los recicladores no tienen un punto formal de acopio en la localidad.'],
  ['S07', 'E14', 1, 'Media', 'Una infraestructura sin uso se deteriora año tras año.'],
  ['S08', 'E15', 4, 'Alta', 'Se rescatan 6 a 9 t/día frente a miles que ingresan: el resto se pierde mientras hay hogares con hambre.'],
  ['S09', 'E16', 3, 'Alta', 'La congestión del corredor retrasa la salida de buses y los servicios de la Troncal Américas.'],
  ['S09', 'E17', 3, 'Media', 'Cuando hay un siniestro, la avenida no tiene vías alternas que absorban el flujo.'],
  ['S10', 'E18', 3, 'Alta', 'La calzada deteriorada frena la carga y las rutas zonales en el último tramo.'],
  ['S10', 'E02', 3, 'Media', 'El mal estado de la vía agrava la mezcla de carga pesada con tránsito local.'],
  ['S11', 'E19', 3, 'Alta', 'Un solo bloqueo obliga a desviar rutas zonales de varios barrios.'],
  ['S11', 'E20', 3, 'Alta', 'Las manifestaciones han llevado a cerrar el Portal Américas y varias estaciones.'],
  ['S12', 'E20', 3, 'Media', 'Las obras del Metro también cierran el portal por tramos.'],
  ['S12', 'E21', 3, 'Media', 'La estación Biblioteca Tintal y el portal operan con capacidad reducida durante las obras.'],
  ['S13', 'E02', 3, 'Severa', 'La carga de un mercado nacional entra por la malla vial de un barrio residencial.'],
  ['S13', 'E16', 3, 'Media', 'El tráfico de carga se suma a la demanda del corredor y aumenta los retrasos.'],
  ['S14', 'E22', 2, 'Crítica', 'La avenida corta el humedal en dos y bloquea el paso de agua y aves entre ambos sectores.'],
  ['S14', 'E23', 2, 'Severa', 'La urbanización y las vías redujeron el humedal a una fracción de su superficie original.'],
  ['S15', 'E24', 2, 'Severa', 'Las invasiones ocupan el suelo que debería almacenar agua.'],
  ['S15', 'E25', 2, 'Alta', 'Los residuos y las conexiones erradas entran al humedal y al drenaje.'],
  ['S16', 'E26', 2, 'Severa', 'La fragmentación por vías y la ocupación reducen el humedal de forma que puede ser irreversible.'],
  ['S16', 'E25', 2, 'Media', 'Escombros y vertimientos llegan al drenaje desde el borde ocupado.'],
  ['S17', 'E27', 2, 'Severa', 'Las conexiones erradas descargan aguas residuales en tuberías de aguas lluvias que van a los ríos.'],
  ['S17', 'E28', 4, 'Crítica', 'El Fucha en La Igualdad está entre los puntos críticos de desbordamiento de la ciudad.'],
  ['S18', 'E29', 4, 'Severa', 'Un barrio asentado sobre el antiguo cauce se inundó gravemente en 1979 y 1980.'],
  ['S18', 'E28', 4, 'Alta', 'Una localidad plana con 35% inundable sufre cuando el río se desborda.'],
  ['S19', 'E30', 4, 'Alta', 'Sin zonas verdes y con suelo sellado aumenta el calor superficial.'],
  ['S20', 'E31', 2, 'Crítica', 'La carga orgánica consume el oxígeno del agua en el borde del humedal.'],
  ['S20', 'E27', 2, 'Alta', 'Los vertimientos del borde de la central siguen hacia los canales y ríos.'],
  // síntomas que se refuerzan
  ['S01', 'S05', 5, 'Alta', 'Cuanto más alimento se mueve, más residuo orgánico se genera en la misma central.'],
  ['S01', 'S13', 5, 'Alta', 'El volumen de la central se traduce directamente en vehículos de carga que entran y salen.'],
  ['S01', 'S08', 5, 'Media', 'El volumen y los excedentes de la central son la base del rescate de alimentos y de las pérdidas.'],
  ['S01', 'S02', 5, 'Alta', 'La central concentra empleo informal y rebusque en su borde.'],
  ['S02', 'S03', 5, 'Alta', 'El reciclaje informal es una de las principales fuentes de ingreso de la economía popular.'],
  ['S02', 'S04', 5, 'Media', 'La informalidad y la brecha digital se refuerzan: sin conexión hay menos empleo calificado.'],
  ['S03', 'S07', 5, 'Media', 'El cierre del centro de reciclaje mantiene la actividad dispersa en bodegas informales.'],
  ['S05', 'S06', 5, 'Media', 'La falta de tecnologías de compostaje afecta tanto a la central como al humedal vecino.'],
  ['S05', 'S20', 5, 'Severa', 'La fracción orgánica sin aprovechar alimenta la carga orgánica de los canales vecinos.'],
  ['S13', 'S10', 5, 'Media', 'El flujo de carga acelera el deterioro de la Calle 38 Sur.'],
  ['S13', 'S09', 5, 'Alta', 'La carga hacia la central se suma a la demanda de la Av. Ciudad de Cali.'],
  ['S09', 'S11', 5, 'Alta', 'Un corredor saturado amplifica el efecto de cualquier bloqueo en Patio Bonito.'],
  ['S11', 'S12', 5, 'Media', 'Los cierres por bloqueos y por obras se suman en las mismas estaciones.'],
  ['S09', 'S14', 5, 'Alta', 'La avenida que se satura es la misma que parte el humedal El Burro.'],
  ['S14', 'S16', 5, 'Media', 'Dos humedales de Kennedy están fragmentados por vías.'],
  ['S15', 'S20', 5, 'Severa', 'La ocupación y la carga orgánica degradan juntas el borde de La Vaca.'],
  ['S15', 'S17', 5, 'Alta', 'Las conexiones erradas en La Vaca se suman a las del Fucha.'],
  ['S17', 'S18', 5, 'Alta', 'El suelo inundable agrava el efecto de las conexiones erradas y los desbordes.'],
  ['S18', 'S19', 5, 'Media', 'Un territorio inundable sin espacio público libre tiene menos dónde absorber agua.']
];

const { S2, E2, ED2 } = require('./datos-sintomas-2');
S.push(...S2); E.push(...E2); ED.push(...ED2);

// ---------- separar nodos muy cercanos para que no se superpongan en el mapa ----------
const pts = []; const nodes = [...S.map(s => ({ s, kind: 'S' })), ...E.map(e => ({ e, kind: 'E' }))];
function ll(n) { return n.kind === 'S' ? n.s[5] : n.e[4]; }
nodes.forEach((n, i) => {
  let [la, lo] = ll(n), k = 0;
  while (pts.some(p => Math.hypot((p[0] - la) * 111000, (p[1] - lo) * 110000) < 45) && k < 40) {
    const a = k * 2.4, r = 0.0004 * (1 + k / 8); la = ll(n)[0] + Math.sin(a) * r; lo = ll(n)[1] + Math.cos(a) * r; k++;
  }
  pts.push([la, lo]); n.latlon = [+la.toFixed(5), +lo.toFixed(5)];
});

// ---------- relaciones entre síntomas: vía de refuerzo + relaciones derivadas de los datos ----------
const layerOf = {}; S.forEach(s => { layerOf[s[0]] = s[2]; });
const OVERRIDE = { 'S30|S49': 8, 'S49|S30': 8 };
function via(a, b) {
  if (OVERRIDE[a + '|' + b]) return OVERRIDE[a + '|' + b];
  const A = layerOf[a], B = layerOf[b];
  if (A === 'Metabolismo urbano' || B === 'Metabolismo urbano') return 6;
  if (A === 'Redes y flujos' || B === 'Redes y flujos') return 7;
  if (A === 'Economía territorial' || B === 'Economía territorial') return 5;
  return 8;
}
ED.forEach(e => { if (e[2] === 5 && e[0][0] === 'S' && e[1][0] === 'S') e[2] = via(e[0], e[1]); });
const pairKey = (a, b) => [a, b].sort().join('|');
const linked = new Set(ED.map(e => pairKey(e[0], e[1])));
const pos = {}; nodes.forEach(n => { pos[n.kind === 'S' ? n.s[0] : n.e[0]] = n.latlon; });
const dist = (a, b) => Math.hypot((pos[a][0] - pos[b][0]) * 111000, (pos[a][1] - pos[b][1]) * 110000);
const ENT = [['IDU', /\bIDU\b/], ['EAAB', /EAAB/], ['UAESP', /UAESP/], ['Secretaría de Movilidad', /Secretaría de Movilidad/], ['TransMilenio', /TransMilenio/], ['SDDE', /SDDE/], ['IPES', /IPES/],
  ['Secretaría de Educación', /Secretaría de Educación/], ['Secretaría de la Mujer', /Secretaría de la Mujer/], ['Secretaría de Ambiente', /Secretaría de Ambiente|\bSDA\b/], ['Secretaría de Salud', /Secretaría de Salud|Subred/],
  ['Empresa Metro de Bogotá', /Empresa Metro|\bMetro\b/]];
const ents = {}; S.forEach(s => { ents[s[0]] = ENT.filter(([, re]) => re.test(s[8])).map(([n]) => n); });
const ids = S.map(s => s[0]);
let nTer = 0, nEnt = 0;
for (let i = 0; i < ids.length; i++) for (let j = i + 1; j < ids.length; j++) {
  const a = ids[i], b = ids[j], k = pairKey(a, b);
  if (linked.has(k)) continue;
  const d = dist(a, b);
  if (d < 600) { ED.push([a, b, 9, d < 250 ? 'Alta' : 'Media', 'Ambos síntomas ocurren a ' + Math.round(d / 10) * 10 + ' m uno del otro: afectan a los mismos vecinos y calles. (Derivada de las coordenadas.)']); linked.add(k); nTer++; continue; }
  const shared = ents[a].filter(x => ents[b].includes(x));
  if (shared.length) { ED.push([a, b, 10, 'Media', 'Una misma entidad tiene competencia sobre ambos síntomas: ' + shared.join(', ') + '. (Derivada de la columna Actores.)']); linked.add(k); nEnt++; }
}
console.log('Relaciones derivadas: territorio ' + nTer + ', entidad ' + nEnt);

(async () => {
  const wb = new ExcelJS.Workbook();
  const HEAD = { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } }, alignment: { vertical: 'middle', wrapText: true } };
  const sheet = (name, cols, rows, widths) => { const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] }); ws.addRow(cols); ws.getRow(1).eachCell(c => Object.assign(c, HEAD)); rows.forEach(r => ws.addRow(r)); widths.forEach((w, i) => ws.getColumn(i + 1).width = w); ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols.length } }; return ws; };
  const ley = wb.addWorksheet('LEEME');
  ['RED DE SÍNTOMAS Y EFECTOS DE KENNEDY — ENFOQUES: ECONOMÍA, INNOVACIÓN, REDES (Y METABOLISMO)', '',
    'Preguntas guía: ¿El territorio genera oportunidades sin agotar recursos ni concentrar beneficios, y puede adaptar su trayectoria? ¿Los flujos y servicios son accesibles y robustos?',
    'Red grande = ' + S.length + ' SÍNTOMAS (hechos observables y georreferenciados). Al presionar un síntoma se abre su sub-red con sus EFECTOS. Al presionar un efecto se ven la imagen 1 y la imagen 2; luego vuelve la red grande y se puede MATERIALIZAR sobre las coordenadas del territorio.',
    '',
    'HOJAS: CAPAS (4 de síntomas + "Efecto en el territorio", que se oculta en la red grande), INTERACCIONES, NODOS, ARISTAS, RED (configuración), CONTROL.',
    'NODOS: columna "Evidencia" = Documentado (con fuente pública), Inferido (verificar en campo) o Dato del grupo (contrastar con fuente oficial). Las cifras contradictorias entre fuentes se indican en el texto.',
    'Latitud/Longitud: OpenStreetMap/Nominatim y Wikipedia para lugares con nombre; los síntomas de alcance local (p. ej. déficit de espacio público) se representan en un barrio afectado y así se indica. Los nodos muy cercanos se separan ~50 m para verlos.',
    'IMÁGENES: columnas "Imagen 1" e "Imagen 2" (URL o ruta assets/...). La columna "Sugerencia de imágenes" dice qué fotografiar. Si están vacías, el visor muestra una ficha con el texto.',
    'GENERAR: node tools/build-red.js datos/red-sintomas-kennedy.xlsx   (luego subir a GitHub)'].forEach((t, i) => { const r = ley.addRow([t]); if (i === 0) r.font = { bold: true, size: 14 }; });
  ley.getColumn(1).width = 160;
  const wsC = sheet('CAPAS', ['Codigo', 'Capa', 'Descripcion en el panel', 'Color (HEX)', 'Icono (Font Awesome)', 'Insignia'], LAYERS, [9, 32, 70, 13, 22, 16]);
  const wsT = sheet('INTERACCIONES', ['Codigo', 'Tipo de relacion', 'Definicion', 'Color (HEX)', 'Icono (Font Awesome)'], TYPES, [9, 38, 80, 13, 26]);
  const rowsN = [];
  nodes.forEach(n => {
    if (n.kind === 'S') { const s = n.s; rowsN.push(['NOD-' + s[0], s[1], s[2], 'Síntoma · ' + s[3], s[10], s[4], s[6], s[7], s[8], s[9], s[11], '', n.latlon[0], n.latlon[1], '', '', s[12], ETIQ[s[0]][1]]); }
    else { const e = n.e; rowsN.push(['NOD-' + e[0], e[1], 'Efecto en el territorio', 'Efecto · ' + e[2], e[7], e[3], e[5], e[6], '', '', e[8], '', n.latlon[0], n.latlon[1], '', '', e[9], ETIQ[e[0]][1]]); }
  });
  const wsN = sheet('NODOS', ['ID', 'Nombre (name)', 'Capa (cat)', 'Categoria / enfoque (sciname)', 'Evidencia', 'Ubicacion (loc)', 'Que se observa / que ocurre (role)', 'Por que es sintoma o por que importa (alert)', 'Actores', 'Pregunta guia', 'Fuente', 'Imagen (img, opcional)', 'Latitud', 'Longitud', 'Imagen 1 (tras la sub-red)', 'Imagen 2', 'Sugerencia de imagenes', 'Icono (Font Awesome)'], rowsN,
    [10, 46, 28, 28, 26, 40, 80, 70, 40, 50, 60, 22, 11, 11, 30, 30, 50, 26]);
  sheet('ARISTAS', ['ID', 'Origen (ID nodo)', 'Destino (ID nodo)', 'Tipo de relacion', 'Detalle del tipo', 'Justificacion causal (rationale)', 'Tension', 'Peso'],
    ED.map((e, i) => ['EDG-' + String(i + 1).padStart(2, '0'), 'NOD-' + e[0], 'NOD-' + e[1], T[e[2]], '', e[4], e[3], W[e[3]]]), [10, 14, 14, 34, 14, 90, 10, 7]);
  const ws = wb.addWorksheet('RED'); ws.addRow(['Clave', 'Valor']); ws.getRow(1).eachCell(c => Object.assign(c, HEAD));
  [['titulo', 'Síntomas y efectos de Kennedy — Economía, innovación y redes'],
    ['marca', 'SÍNTOMAS Y EFECTOS · KENNEDY'],
    ['capa_efectos', 'Efecto en el territorio'],
    ['capa_clave', 'Metabolismo urbano'],
    ['fuentes', 'Alcaldía de Bogotá y Alcaldía Local de Kennedy|Secretaría de la Mujer (EM 2017)|Secretaría de Educación (Kennedy 2022)|Prensa y entidades distritales (ver Fuente de cada nodo)'],
    ['faq1_pregunta', '¿Qué es un síntoma y qué es un efecto?'],
    ['faq1_respuesta', 'Un síntoma es un hecho observable y localizado en Kennedy (p. ej. la Calle 38 Sur deteriorada). Un efecto es lo que ese hecho provoca en el territorio (p. ej. rutas zonales desviadas). Al presionar un síntoma aparecen sus efectos.'],
    ['faq2_pregunta', '¿Cómo se relacionan economía, innovación y redes?'],
    ['faq2_respuesta', 'La economía concentra el flujo en pocos nodos (Corabastos); la innovación debería transformar sus residuos y su conocimiento, pero la brecha digital y los centros cerrados lo impiden; las redes de transporte y de agua no son robustas ante bloqueos, obras o lluvias.'],
    ['faq3_pregunta', '¿Qué datos son verificados?'],
    ['faq3_respuesta', 'Cada nodo indica su evidencia: Documentado (fuente pública), Inferido (por verificar en campo) o Dato del grupo (por contrastar con una fuente oficial). Las cifras que varían entre fuentes se señalan en el texto.'],
    ['faq4_pregunta', '¿Cómo veo todo en el territorio?'],
    ['faq4_respuesta', 'Después de las imágenes vuelve la red grande; pulsa MATERIALIZAR y los síntomas y sus efectos aparecen sobre las coordenadas de Kennedy.']].forEach(r => ws.addRow(r));
  ws.getColumn(1).width = 20; ws.getColumn(2).width = 130;
  const ctl = wb.addWorksheet('CONTROL'); ctl.addRow(['Prueba', 'Resultado', 'Estado']); ctl.getRow(1).eachCell(c => Object.assign(c, HEAD));
  const tests = [
    ['Nodos con ID duplicado', 'SUMPRODUCT((COUNTIF(NODOS!A2:A500,NODOS!A2:A500)>1)*(NODOS!A2:A500<>""))'],
    ['Nodos con capa no valida', 'SUMPRODUCT((NODOS!A2:A500<>"")*(COUNTIF(CAPAS!B2:B30,NODOS!C2:C500)=0))'],
    ['Aristas con origen inexistente', 'SUMPRODUCT((ARISTAS!A2:A500<>"")*(COUNTIF(NODOS!A2:A500,ARISTAS!B2:B500)=0))'],
    ['Aristas con destino inexistente', 'SUMPRODUCT((ARISTAS!A2:A500<>"")*(COUNTIF(NODOS!A2:A500,ARISTAS!C2:C500)=0))'],
    ['Aristas con tipo fuera de catalogo', 'SUMPRODUCT((ARISTAS!A2:A500<>"")*(COUNTIF(INTERACCIONES!B2:B30,ARISTAS!D2:D500)=0))'],
    ['Autoaristas', 'SUMPRODUCT((ARISTAS!A2:A500<>"")*(ARISTAS!B2:B500=ARISTAS!C2:C500))'],
    ['Nodos aislados (grado 0)', 'SUMPRODUCT((NODOS!A2:A500<>"")*(COUNTIF(ARISTAS!B2:B500,NODOS!A2:A500)+COUNTIF(ARISTAS!C2:C500,NODOS!A2:A500)=0))']];
  tests.forEach((t, i) => { const r = i + 2; ctl.addRow([t[0], { formula: t[1] }, { formula: `IF(B${r}=0,"OK","REVISAR")` }]); });
  ctl.addRow(['ESTADO GLOBAL', '', { formula: `IF(COUNTIF(C2:C${tests.length + 1},"OK")=${tests.length},"RED LISTA","REVISAR")` }]);
  ctl.addRow(['Total nodos', { formula: 'COUNTA(NODOS!A2:A500)' }]); ctl.addRow(['Total aristas', { formula: 'COUNTA(ARISTAS!A2:A500)' }]);
  ctl.getColumn(1).width = 44; ctl.getColumn(2).width = 14; ctl.getColumn(3).width = 16;
  [[wsC, 4, LAYERS.length], [wsT, 4, TYPES.length]].forEach(([w, col, n]) => { for (let r = 2; r <= n + 1; r++) { const c = w.getCell(r, col); c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + String(c.value).replace('#', '') } }; c.font = { bold: true }; } });
  await wb.xlsx.writeFile(OUT);
  console.log(`OK: ${S.length} síntomas, ${E.length} efectos, ${ED.length} relaciones -> ${OUT}`);
})();
