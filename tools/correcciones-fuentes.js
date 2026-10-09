// Fuentes primarias (oficiales o académicas) añadidas a los nodos que solo tenían prensa o estaban en "verificar".
// Se aplica sobre S (síntomas) y E (efectos) antes de generar el Excel. Cada fuente fue localizada en su sitio oficial/repositorio.
module.exports = function (S, E) {
  const byId = id => S.find(r => r[0] === id) || E.find(r => r[0] === id);
  const isS = id => id[0] === 'S';
  const SRC = r => (r[0][0] === 'S' ? 11 : 8), ROLE = r => (r[0][0] === 'S' ? 6 : 5), WHY = r => (r[0][0] === 'S' ? 7 : 6), EV = r => (r[0][0] === 'S' ? 10 : 7);
  const add = (id, extra) => { const r = byId(id); if (!r) throw new Error('nodo ' + id); r[SRC(r)] = r[SRC(r)] + '; ' + extra; };
  const role = (id, t) => { const r = byId(id); r[ROLE(r)] = t; };
  const why = (id, t) => { const r = byId(id); r[WHY(r)] = t; };
  const name = (id, t) => { const r = byId(id); r[1] = t; };
  const ev = (id, t) => { const r = byId(id); r[EV(r)] = t; };

  // ---- referencias reutilizables ----
  const TM_STATS = 'TransMilenio S.A., Estadísticas de oferta y demanda del SITP, diciembre 2025 (transmilenio.gov.co)';
  const TM_CIERRE = 'TransMilenio S.A., Novedades operativas en el Portal Américas y estación Biblioteca Tintal por obras Metro (transmilenio.gov.co/publicaciones/154711/novedades-operativas-en-el-portal-americas-y-estacion-biblioteca-tintal-por-obras-metro)';
  const BLOQ = 'Alcaldía de Bogotá, "Más de 30.000 personas afectadas por bloqueos a TransMilenio" (bogota.gov.co/mi-ciudad/movilidad/mas-de-30000-personas-afectadas-por-bloqueos-transmilenio)';
  const MOV_SINIES = 'Secretaría Distrital de Movilidad, Observatorio de Movilidad, Análisis de la siniestralidad vial en Bogotá (observatorio.movilidadbogota.gov.co/actualidad/siniestralidad-vial-en-bogota) y Anuario de Siniestralidad Vial 2024 (movilidadbogota.gov.co)';
  const BCV_KEN = 'Bogotá Cómo Vamos, Una mirada a la siniestralidad vial en Kennedy (bogotacomovamos.org/siniestralidad-vial-en-kennedy/)';
  const UAESP51 = 'UAESP, 51 toneladas de residuos se recogieron durante operativo de limpieza en Kennedy (uaesp.gov.co/noticias/51-toneladas-residuos-se-recogieron-durante-operativo-limpieza-kennedy)';
  const ALC22K = 'Alcaldía Local de Kennedy, más de 22.000 m² de espacio público recuperados (gobiernobogota.gov.co/noticias/alcaldia-local-de-kennedy-ha-recuperado-mas-de-22-mil-metros-cuadrados-de-espacio-publico-afectados-por-residuos)';
  const ANDAL = 'Alcaldía Local de Kennedy, recupera 1.500 m² y retira 38 toneladas en el barrio Andalucía (gobiernobogota.gov.co/noticias/kennedy-recupera-1500-m2-espacio-publico)';
  const EAAB_TINTAL = 'Alcaldía de Bogotá y EAAB, daños imprevistos de tubería en Kennedy: Av. Américas con Carrera 78 (diciembre de 2025) y barrio Tintal (abril de 2023) (bogota.gov.co/mi-ciudad/habitat/acueducto-de-bogota-atiende-dano-imprevisto-en-localidad-de-kennedy; bogota.gov.co/mi-ciudad/habitat/acueducto-de-bogota-atiende-contingencia-en-tuberia-en-kennedy)';
  const IDIGER = 'IDIGER, Inundaciones en Bogotá: causas y zonas de riesgo (idiger.gov.co/escenarios-de-riesgo/inundacion/en-bogota)';
  const ASIS = 'Secretaría Distrital de Salud, Observatorio de Salud, ASIS Localidad Kennedy 2024 (saludata.saludcapital.gov.co/osb/wp-content/uploads/2026/01/8.-ASIS_Kennedy_2024.pdf)';
  const MORCI = 'Alcaldía Local de Kennedy, Sellada fábrica de morcilla insalubre en Kennedy (gobiernobogota.gov.co/noticias/sellada-fabrica-morcilla-insalubre-Kennedy-IVC)';
  const BICI = 'Secretaría Distrital de Movilidad, caracterización y censo del bicitaxismo (bogota.gov.co/mi-ciudad/movilidad/bicitaxismo-en-bogota)';
  const CCB = 'Cámara de Comercio de Bogotá (2020), Perfil económico y empresarial de las localidades de Bogotá (repositoriocdim.esap.edu.co)';

  // ---- síntomas ----
  add('S09', 'Alcaldía de Bogotá, avance de la Línea 1 y obras en el corredor (bogota.gov.co/mi-ciudad/movilidad)');
  add('S11', BLOQ);
  add('S12', TM_CIERRE + '; Alcaldía de Bogotá, Se posterga cierre temporal en Portal Américas y estación Tintal: una dovela no cumplió los controles técnicos en el viaducto de la Av. Ciudad de Cali con Villavicencio (bogota.gov.co/mi-ciudad/movilidad/movilidad-en-bogota-portal-americas-y-biblioteca-tintal-por-obras-2025)');
  role('S40', 'Según TransMilenio, el Portal Américas es el portal con más validaciones diarias del sistema: 77.167 en diciembre de 2024, 74.774 en marzo de 2025 y 70.315 en diciembre de 2025 (12.115 en hora pico). En 2023 vecinos describieron filas de ingreso de más de 30 minutos; en noviembre de 2024 hubo congestión en su salida y en abril de 2026 un accidente obligó a desviar la flota troncal.');
  add('S40', TM_STATS);
  name('S40', 'El Portal Américas concentra unas 70.000 validaciones diarias y es el portal más usado del sistema');
  role('S41', 'TransMilenio y la Alcaldía han reportado bloqueos recurrentes en la Troncal Américas y el Portal Américas: en un caso más de 30.000 usuarios afectados y las estaciones Biblioteca Tintal y Patio Bonito sin servicio; en un corte de diciembre de 2024, 10.199 usuarios afectados. En agosto de 2026 un bloqueo dejó más de 20.000 usuarios afectados (20.074) y cinco estaciones sin servicio (reporte de prensa).');
  add('S41', BLOQ);
  name('S41', 'Los bloqueos en la Troncal Américas afectan a decenas de miles de usuarios y dejan estaciones sin servicio');
  role('S44', 'La Secretaría de Movilidad censó 3.054 bicitaxis en Bogotá en 2013 y 4.646 en 2019 (+52%); el servicio solo está autorizado en modalidad no motorizada (Resolución 3256 de 2018 del Ministerio de Transporte). La prensa reporta más de 5.000 bicitaxistas en sectores como Patio Bonito y El Tintal y un vehículo adaptado con unas nueve personas de pie y sin medidas de seguridad.');
  add('S44', BICI + '; Ministerio de Transporte, Resolución 3256 de 2018 (transporte en triciclos no motorizados)');
  role('S46', 'La Secretaría de Movilidad ubica a Kennedy entre las localidades con más víctimas fatales por siniestros viales (Anuario 2024) y registra un promedio de 70 fallecidos al año entre 2015 y 2022; en los primeros cinco meses de 2023 hubo 34 víctimas fatales en la localidad. La prensa reportó muertes de ciclistas en la Av. Boyacá: tractocamión en el cruce con la Av. Américas (febrero de 2023), volqueta en la Calle 44 Sur (junio de 2024) y camión en la Calle 7C (octubre de 2025).');
  add('S46', MOV_SINIES + '; ' + BCV_KEN);
  role('S47', 'La Secretaría de Movilidad ubica a Kennedy entre las localidades con más víctimas fatales por siniestros viales y los motociclistas fueron el 47,3% de los fallecidos de Bogotá en 2024. La prensa reportó que una vía en mal estado en Kennedy cobró la vida de un motociclista y que un accidente en la Av. Primero de Mayo se atribuyó a una obra sin terminar y a un carril en mal estado (2023).');
  add('S47', MOV_SINIES + '; Alcaldía de Bogotá, mejoramiento de la ciclorruta de la Av. Primero de Mayo (84 segmentos, 1,83 km; bogota.gov.co/asi-vamos/obras/mejoramiento-de-la-malla-vial-en-la-ciclorruta-del-hospital-de-kennedy)');
  add('S50', ALC22K + ' (564 bodegas de reciclaje intervenidas y 18 puntos críticos atendidos)');
  role('S51', 'La UAESP reportó 51 toneladas de residuos recogidas en un operativo en Patio Bonito, El Jazmín y Bellavista (febrero de 2023), con 5 bodegas de reciclaje selladas. En septiembre de 2025 la prensa reportó entre 45 y 55 toneladas retiradas de un espacio de 12.000 m² en Jazmín Occidental, con bodegas informales desmanteladas cerca de obras del Metro (las cifras de las fuentes no coinciden).');
  add('S51', UAESP51 + '; ' + ALC22K);
  add('S52', ANDAL);
  add('S61', ALC22K + ' (564 bodegas de reciclaje intervenidas); ' + ANDAL);
  role('S56', 'La EAAB ha atendido daños imprevistos de tubería que suspendieron el agua en Kennedy: noviembre de 2021 (tubería de 24 pulgadas, entre la calle 6 y la 16C, barrios Tintal, Castilla y Villa Mejía), abril de 2023 (Av. Calle 6 con Carrera 89B, barrio Tintal) y diciembre de 2025 (Av. Américas con Carrera 78: desde la calle 9 hasta la 43 Sur, incluidos Tintalá y Tintalito, con Acualínea 116 y carrotanques).');
  add('S56', EAAB_TINTAL);
  add('S60', ASIS + ' (la inundación es el principal riesgo del plan local de emergencias de Kennedy); no se encontró comunicado oficial del evento en el hospital');
  ev('S60', 'Prensa; contexto de riesgo documentado (verificar el evento con la Subred Sur Occidente)');
  add('S18', IDIGER + '; ' + ASIS + ' (tres zonas de inundación en Kennedy: microcuenca del Fucha, entre el canal Cundinamarca y el río Bogotá, y ronda del Tunjuelo); Decreto 555 de 2021, mapas de amenaza por inundación y por encharcamiento (CU-2.2.14)');
  add('S37', MORCI);
  add('S22', CCB);
  add('S04', 'Alcaldía de Bogotá, programa Conexión Social con ETB (bogota.gov.co)');

  // ---- efectos ----
  add('E01', 'Alcaldía de Bogotá, Paro de camioneros en Bogotá ya genera alzas en alimentos y pérdidas económicas (bogota.gov.co/mi-ciudad/desarrollo-economico/paro-de-camioneros-en-bogota-alzas-en-alimentos-y-perdidas-economicas)');
  add('E07', 'Secretaría Distrital de la Mujer, Diagnóstico local Kennedy 2020 (conectividad de los hogares con jefatura femenina)');
  add('E20', TM_CIERRE);
  add('E33', CCB);
  add('E51', TM_STATS);
  add('E52', BLOQ);
  add('E53', 'Alcaldía de Bogotá, cierre parcial de la estación Patio Bonito por obras de la Av. Ciudad de Cali (bogota.gov.co/en/node/177989)');
  add('E55', BICI);
  add('E56', 'Secretaría Distrital de Movilidad, Resolución 3256 de 2018 del Ministerio de Transporte sobre bicitaxis (contexto normativo)');
  add('E57', MOV_SINIES);
  add('E58', MOV_SINIES);
  add('E64', UAESP51);
  add('E65', ANDAL);
  add('E70', EAAB_TINTAL);
  add('E74', ASIS);
  add('E75', MORCI);
  add('E29', IDIGER + ' (historial de inundaciones; en diciembre de 2011 el reflujo del Fucha inundó las partes bajas de Kennedy y Bosa)');
  name('E30', 'Calor superficial de hasta 33 °C con 78% de suelo sellado (dato del grupo; por contrastar)');
  role('E30', 'La memoria del grupo registra 78,4% de suelo cubierto por asfalto y concreto y temperaturas superficiales de 28,5 a 33,2 °C. La nota técnica de la SDA sobre el efecto isla de calor (temperatura del aire, 45 estaciones, 2008–2018) ubica el mayor efecto en Bosa, Ciudad Bolívar, Los Mártires, Tunjuelito y Rafael Uribe, y Kennedy no figura entre ellas: la temperatura superficial y la del aire no son comparables.');
  add('E30', 'Secretaría Distrital de Ambiente, nota técnica sobre arbolado urbano y efecto isla de calor (oab.ambientebogota.gov.co/descargar/29103); Aragón et al. (2020), Geographicalia, temperatura superficial con Landsat 8 en la Sabana de Bogotá (papiro.unizar.es/ojs/index.php/geographicalia/article/view/4571)');
};
