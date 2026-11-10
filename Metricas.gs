// ============================================================
//  Eg_Metricas — Sistema de Métricas Redes Sociales
//  Google Apps Script | Grupo LPG
//  Plataformas: YouTube · Facebook · Instagram · TikTok · Spotify
// ============================================================

// ── CONFIGURACIÓN ───────────────────────────────────────────
const CONFIG = {
  YOUTUBE_API_KEY:        '',   // Google Cloud Console → YouTube Data API v3
  // Tokens por página de Facebook (pega el access_token de cada una)
  META_TOKENS: {
    '119536778059759': '',   // El Gráfico
    '107594569282760': '',   // LPG La Prensa Gráfica
  },
  TIKTOK_ACCESS_TOKEN:    '',   // TikTok for Developers → Display API token
  SPOTIFY_CLIENT_ID:      '',   // Spotify for Developers → App credentials
  SPOTIFY_CLIENT_SECRET:  '',

  // ── Columnas de ENTRADA (estructura existente del Sheet) ──
  COL_LINK:         1,   // A - URL del contenido
  COL_FECHA:        2,   // B - Fecha
  COL_PROGRAMA:     3,   // C - Programa
  COL_PLATAFORMA:   4,   // D - Red (YouTube, Facebook, etc.)
  COL_TIPO:         5,   // E - Tipo (Video, Audio, etc.)
  COL_TIPO_PAUTA:   6,   // F - Tipo de pauta
  COL_ESTADO:       7,   // G - Estado
  COL_MARCA:        8,   // H - Marca

  // ── Columnas de SALIDA (métricas — se escriben al final) ──
  COL_VIEWS:        9,
  COL_REACH:        10,
  COL_LIKES:        11,
  COL_COMENTARIOS:  12,
  COL_SHARES:       13,
  COL_ULTIMA_ACT:   14,
  COL_IG_MEDIA_ID:  15,  // Media ID de Instagram (se llena con indexarInstagram)

  FILA_INICIO: 2,   // Primera fila de datos (la 1 es encabezado)
  SPREADSHEET_ID: '',
};


// ── MENÚ PERSONALIZADO ──────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Métricas SIMAN')
    .addItem('Actualizar hoja activa', 'actualizarHojaActiva')
    .addItem('Actualizar todas las hojas', 'actualizarTodasLasHojas')
    .addSeparator()
    .addItem('Configurar encabezados', 'configurarEncabezados')
    .addToUi();
}


// ── RUNNERS ─────────────────────────────────────────────────
function actualizarHojaActiva() {
  const hoja = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getActiveSheet();
  procesarHoja(hoja);
}

function actualizarTodasLasHojas() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  ss.getSheets().forEach(hoja => procesarHoja(hoja));
}

// Procesa filas en bloques — útil si el Sheet tiene muchas filas
// Cambia BLOQUE_INICIO y BLOQUE_FIN según el rango que quieras procesar
// Bloques de 100 filas para cubrir 969 filas
function actualizarBloque()  { procesarHoja(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getActiveSheet(), 2,   101); }
function actualizarBloque2() { procesarHoja(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getActiveSheet(), 102, 201); }
function actualizarBloque3() { procesarHoja(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getActiveSheet(), 202, 301); }
function actualizarBloque4() { procesarHoja(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getActiveSheet(), 302, 401); }
function actualizarBloque5() { procesarHoja(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getActiveSheet(), 402, 501); }
function actualizarBloque6() { procesarHoja(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getActiveSheet(), 502, 601); }
function actualizarBloque7() { procesarHoja(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getActiveSheet(), 602, 701); }
function actualizarBloque8() { procesarHoja(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getActiveSheet(), 702, 801); }
function actualizarBloque9() { procesarHoja(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getActiveSheet(), 802, 901); }
function actualizarBloque10(){ procesarHoja(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getActiveSheet(), 902, 970); }


// ── PROCESADOR DE HOJA ──────────────────────────────────────
function procesarHoja(hoja, filaInicio, filaFin) {
  const datos = hoja.getDataRange().getValues();
  const inicio = (filaInicio || CONFIG.FILA_INICIO) - 1;
  const fin    = filaFin ? Math.min(filaFin - 1, datos.length - 1) : datos.length - 1;

  for (let i = inicio; i <= fin; i++) {
    const fila = datos[i];
    const link = fila[CONFIG.COL_LINK - 1];
    if (!link || !link.toString().startsWith('http')) continue;

    const plataforma    = detectarPlataforma(link);

    // Solo procesar plataformas soportadas
    if (!['YouTube', 'Facebook', 'Instagram', 'TikTok'].includes(plataforma)) continue;

    const tipo          = (fila[CONFIG.COL_TIPO - 1] || '').toString().toLowerCase().trim()
                          || detectarTipo(plataforma, link);
    const igMediaId     = (fila[CONFIG.COL_IG_MEDIA_ID - 1] || '').toString().trim();
    const metricas      = obtenerMetricas(plataforma, link, tipo, igMediaId);

    if (!metricas) continue;

    const rango = hoja.getRange(i + 1, CONFIG.COL_VIEWS, 1, 6);
    rango.setValues([[
      metricas.views       ?? '',
      metricas.reach       ?? '',
      metricas.likes       ?? '',
      metricas.comentarios ?? '',
      metricas.shares      ?? '',
      new Date()
    ]]);

    SpreadsheetApp.flush();
    Utilities.sleep(300); // evitar quota exceeded
  }

  Logger.log(`✅ "${hoja.getName()}" actualizada`);
}


// ── DETECTOR DE PLATAFORMA ──────────────────────────────────
function detectarPlataforma(url) {
  url = url.toLowerCase();
  if (url.includes('youtube.com') || url.includes('youtu.be')) return 'YouTube';
  if (url.includes('facebook.com') || url.includes('fb.watch'))  return 'Facebook';
  if (url.includes('instagram.com'))                              return 'Instagram';
  if (url.includes('tiktok.com'))                                 return 'TikTok';
  if (url.includes('spotify.com'))                                return 'Spotify';
  return 'Desconocido';
}


// ── DETECTOR DE TIPO (fallback si columna vacía) ────────────
function detectarTipo(plataforma, url) {
  if (plataforma === 'YouTube') {
    if (url.includes('/shorts/'))  return 'short';
    if (url.includes('/live/'))    return 'live';
    return 'video';
  }
  if (plataforma === 'Facebook') {
    if (url.includes('/reel/'))    return 'reel';
    if (url.includes('/videos/'))  return 'video';
    return 'post';
  }
  if (plataforma === 'Instagram') {
    if (url.includes('/reel/'))    return 'reel';
    return 'post';
  }
  if (plataforma === 'Spotify') {
    if (url.includes('/episode/')) return 'episode';
    return 'track';
  }
  return 'video';
}


// ── ROUTER DE MÉTRICAS ──────────────────────────────────────
function obtenerMetricas(plataforma, url, tipo, igMediaId) {
  try {
    switch (plataforma) {
      case 'YouTube':   return metricasYouTube(url);
      case 'Facebook':  return metricasFacebook(url, tipo);
      case 'Instagram': return metricasInstagram(url, tipo, igMediaId);
      case 'TikTok':    return metricasTikTok(url);
      case 'Spotify':   return metricasSpotify(url);
      default:          return null;
    }
  } catch (e) {
    Logger.log(`Error [${plataforma}] ${url}: ${e.message}`);
    return null;
  }
}


// ── YOUTUBE ─────────────────────────────────────────────────
// Soporta: /watch?v=ID  /live/ID  /shorts/ID  youtu.be/ID
function parsearIdYouTube(url) {
  const patrones = [
    /[?&]v=([a-zA-Z0-9_-]{11})/,
    /youtu\.be\/([a-zA-Z0-9_-]{11})/,
    /\/(?:live|shorts|embed)\/([a-zA-Z0-9_-]{11})/,
  ];
  for (const p of patrones) {
    const m = url.match(p);
    if (m) return m[1];
  }
  return null;
}

function metricasYouTube(url) {
  const id = parsearIdYouTube(url);
  if (!id) return null;

  const endpoint = `https://www.googleapis.com/youtube/v3/videos`
    + `?part=statistics&id=${id}&key=${CONFIG.YOUTUBE_API_KEY}`;
  const resp = UrlFetchApp.fetch(endpoint, { muteHttpExceptions: true });
  const json = JSON.parse(resp.getContentText());

  if (!json.items || json.items.length === 0) return null;
  const s = json.items[0].statistics;

  return {
    views:       parseInt(s.viewCount      || 0),
    likes:       parseInt(s.likeCount      || 0),
    comentarios: parseInt(s.commentCount   || 0),
    shares:      null,   // YouTube no expone shares por API pública
    reach:       null,
  };
}


// ── FACEBOOK ────────────────────────────────────────────────
// Soporta: /reel/ID  /videos/ID  /watch/?v=ID
function parsearIdFacebook(url) {
  const patrones = [
    /\/reel\/(\d+)/,
    /\/videos\/(\d+)/,
    /[?&]v=(\d+)/,
    /\/(\d{10,})\/?$/,
  ];
  for (const p of patrones) {
    const m = url.match(p);
    if (m) return m[1];
  }
  return null;
}

// Detecta qué página publicó el contenido y devuelve su token
function resolverTokenFacebook(contentId) {
  for (const [pageId, token] of Object.entries(CONFIG.META_TOKENS)) {
    if (!token) continue;
    const check = `https://graph.facebook.com/v19.0/${contentId}?fields=id&access_token=${token}`;
    try {
      const r = JSON.parse(UrlFetchApp.fetch(check, { muteHttpExceptions: true }).getContentText());
      if (r.id && !r.error) return token;
    } catch(e) {}
  }
  return Object.values(CONFIG.META_TOKENS).find(t => t) || '';
}

function metricasFacebook(url, tipo) {
  const id = parsearIdFacebook(url);
  if (!id) return null;

  const token = resolverTokenFacebook(id);
  if (!token) { Logger.log('Facebook: no hay token configurado'); return null; }

  let campos;
  if (tipo === 'post') {
    campos = 'likes.summary(true),comments.summary(true),shares';
  } else {
    campos = 'likes.summary(true),comments.summary(true),video_insights';
  }

  const endpoint = `https://graph.facebook.com/v19.0/${id}`
    + `?fields=${campos}&access_token=${token}`;
  const resp = UrlFetchApp.fetch(endpoint, { muteHttpExceptions: true });
  const json = JSON.parse(resp.getContentText());

  if (json.error) {
    Logger.log('Facebook error: ' + json.error.message);
    return null;
  }

  let views = null;
  if (tipo !== 'post' && json.video_insights && json.video_insights.data) {
    const vMetric = json.video_insights.data.find(d => d.name === 'post_video_views');
    if (vMetric) views = vMetric.values?.[0]?.value ?? null;
  }

  return {
    views:       views,
    reach:       null,
    likes:       json.likes?.summary?.total_count    ?? null,
    comentarios: json.comments?.summary?.total_count ?? null,
    shares:      json.shares?.count                  ?? null,
  };
}


// ── INSTAGRAM ───────────────────────────────────────────────
// IDs de cuentas Instagram Business vinculadas a páginas FB
const IG_ACCOUNTS = {
  '119536778059759': '17841400861833885',  // El Gráfico → @elgraficionado
  '107594569282760': '17841400180120835',  // LPG        → @laprensagrafica
};

function parsearIdInstagram(url) {
  const m = url.match(/\/(p|reel|tv)\/([A-Za-z0-9_-]+)/);
  return m ? m[2] : null;
}

// Caché global de shortcode → {mediaId, token}
let _igCache = null;

function construirCacheInstagram() {
  if (_igCache) return _igCache;
  _igCache = {};
  for (const [pageId, igId] of Object.entries(IG_ACCOUNTS)) {
    const token = CONFIG.META_TOKENS[pageId];
    if (!token) continue;
    let url = `https://graph.facebook.com/v19.0/${igId}/media`
      + `?fields=id,shortcode&limit=100&access_token=${token}`;
    let paginas = 0;
    while (url && paginas < 50) {
      try {
        const r = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
        if (!r.data) break;
        r.data.forEach(m => { _igCache[m.shortcode] = { mediaId: m.id, token }; });
        url = r.paging && r.paging.next ? r.paging.next : null;
        if (url) Utilities.sleep(150);
        paginas++;
      } catch(e) {
        Logger.log('Instagram cache error: ' + e.message);
        break;
      }
    }
  }
  Logger.log(`Instagram caché: ${Object.keys(_igCache).length} posts indexados`);
  return _igCache;
}

function resolverMediaInstagram(shortcode) {
  const cache = construirCacheInstagram();
  return cache[shortcode] || null;
}

function metricasInstagramPorId(mediaId, token, tipo) {
  // video_views solo disponible en reels — para posts genera error #100
  const campos = 'like_count,comments_count,media_product_type';
  const endpoint = `https://graph.facebook.com/v19.0/${mediaId}`
    + `?fields=${campos}&access_token=${token}`;
  const resp = UrlFetchApp.fetch(endpoint, { muteHttpExceptions: true });
  const json = JSON.parse(resp.getContentText());
  if (json.error) { Logger.log('Instagram error: ' + json.error.message); return null; }
  return {
    views:       json.video_views   ?? null,
    reach:       null,
    likes:       json.like_count     ?? null,
    comentarios: json.comments_count ?? null,
    shares:      null,
  };
}

function metricasInstagram(url, tipo, mediaIdGuardado) {
  // Si ya tenemos el media ID guardado en el Sheet, úsalo directamente
  if (mediaIdGuardado && mediaIdGuardado !== 'NOT_FOUND') {
    const token = Object.values(CONFIG.META_TOKENS).find(t => t) || '';
    return metricasInstagramPorId(mediaIdGuardado, token, tipo);
  }

  const shortcode = parsearIdInstagram(url);
  if (!shortcode) return null;

  const resolved = resolverMediaInstagram(shortcode);
  if (!resolved) { Logger.log('Instagram: media no encontrado para shortcode ' + shortcode); return null; }

  return metricasInstagramPorId(resolved.mediaId, resolved.token, tipo);
}


// ── TIKTOK ──────────────────────────────────────────────────
// Soporta: /@usuario/video/ID
function parsearIdTikTok(url) {
  const m = url.match(/\/video\/(\d+)/);
  return m ? m[1] : null;
}

// Cuenta de TikTok autorizada (la API solo entrega videos de la cuenta que autoriza)
const TIKTOK_CUENTA = 'campuslife.sv';

function metricasTikTok(url) {
  const id = parsearIdTikTok(url);
  if (!id) return null;
  if (url.toLowerCase().indexOf('@' + TIKTOK_CUENTA) === -1) {
    Logger.log('TikTok: omitido (no es de @' + TIKTOK_CUENTA + ') ' + url);
    return null;
  }
  const token = obtenerTokenTikTok();
  if (!token) { Logger.log('TikTok: sin token. Ejecuta autorizarTikTok().'); return null; }

  const resp = UrlFetchApp.fetch(
    'https://open.tiktokapis.com/v2/video/query/?fields=id,view_count,like_count,comment_count,share_count',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ filters: { video_ids: [id] } }),
      muteHttpExceptions: true,
    });
  const json = JSON.parse(resp.getContentText());
  if ((json.error && json.error.code !== 'ok') || !json.data || !json.data.videos || !json.data.videos.length) {
    Logger.log('TikTok ' + id + ': ' + resp.getContentText());
    return null;
  }
  const v = json.data.videos[0];
  return {
    views: v.view_count ?? null,
    reach: null, // la Display API no entrega alcance
    likes: v.like_count ?? null,
    comentarios: v.comment_count ?? null,
    shares: v.share_count ?? null,
  };
}

// Token: se renueva solo con el refresh_token (guardado en Script Properties)
function obtenerTokenTikTok() {
  const props = PropertiesService.getScriptProperties();
  const refresh = props.getProperty('TIKTOK_REFRESH_TOKEN');
  if (!refresh) return CONFIG.TIKTOK_ACCESS_TOKEN || null;
  const exp = Number(props.getProperty('TIKTOK_TOKEN_EXP') || 0);
  const actual = props.getProperty('TIKTOK_ACCESS_TOKEN');
  if (actual && Date.now() < exp - 60000) return actual;

  const r = UrlFetchApp.fetch('https://open.tiktokapis.com/v2/oauth/token/', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      client_key: props.getProperty('TIKTOK_CLIENT_KEY'),
      client_secret: props.getProperty('TIKTOK_CLIENT_SECRET'),
      grant_type: 'refresh_token',
      refresh_token: refresh,
    },
    muteHttpExceptions: true,
  });
  const j = JSON.parse(r.getContentText());
  if (!j.access_token) { Logger.log('TikTok refresh falló: ' + r.getContentText()); return null; }
  props.setProperty('TIKTOK_ACCESS_TOKEN', j.access_token);
  props.setProperty('TIKTOK_REFRESH_TOKEN', j.refresh_token || refresh);
  props.setProperty('TIKTOK_TOKEN_EXP', String(Date.now() + j.expires_in * 1000));
  return j.access_token;
}

// Guarda credenciales y code de la autorización (ejecutar una sola vez)
// 1) Rellena los 4 valores, 2) ejecuta esta función.
function autorizarTikTok() {
  const CLIENT_KEY = '', CLIENT_SECRET = '', CODE = '', REDIRECT_URI = '';
  if (!CLIENT_KEY || !CODE) { Logger.log('Completa CLIENT_KEY, CLIENT_SECRET, CODE y REDIRECT_URI.'); return; }
  const r = UrlFetchApp.fetch('https://open.tiktokapis.com/v2/oauth/token/', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: { client_key: CLIENT_KEY, client_secret: CLIENT_SECRET, code: CODE,
               grant_type: 'authorization_code', redirect_uri: REDIRECT_URI },
    muteHttpExceptions: true,
  });
  const j = JSON.parse(r.getContentText());
  if (!j.access_token) { Logger.log('Error: ' + r.getContentText()); return; }
  const p = PropertiesService.getScriptProperties();
  p.setProperties({
    TIKTOK_CLIENT_KEY: CLIENT_KEY, TIKTOK_CLIENT_SECRET: CLIENT_SECRET,
    TIKTOK_ACCESS_TOKEN: j.access_token, TIKTOK_REFRESH_TOKEN: j.refresh_token,
    TIKTOK_TOKEN_EXP: String(Date.now() + j.expires_in * 1000),
  });
  Logger.log('TikTok autorizado. Cuenta: ' + j.open_id);
}

// Prueba: lista los últimos videos de la cuenta autorizada con sus métricas
function verMisDatosTikTok() {
  const token = obtenerTokenTikTok();
  if (!token) { Logger.log('Sin token. Ejecuta autorizarTikTok().'); return; }
  const r = UrlFetchApp.fetch(
    'https://open.tiktokapis.com/v2/video/list/?fields=id,title,create_time,view_count,like_count,comment_count,share_count',
    { method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ max_count: 20 }), muteHttpExceptions: true });
  const j = JSON.parse(r.getContentText());
  (j.data && j.data.videos || []).forEach(v =>
    Logger.log(v.id + ' | ' + (v.title||'').slice(0,40) + ' | 👁 ' + v.view_count + ' ❤ ' + v.like_count + ' 💬 ' + v.comment_count + ' ↗ ' + v.share_count));
  if (!j.data) Logger.log(r.getContentText());
}


// ── SPOTIFY ─────────────────────────────────────────────────
// Soporta: /episode/ID  /track/ID  /show/ID
// Nota: Spotify no expone play counts públicos.
// Para podcasts con Spotify for Podcasters API sí hay acceso.
function parsearSpotify(url) {
  const m = url.match(/spotify\.com\/(track|episode|show)\/([A-Za-z0-9]+)/);
  return m ? { tipo: m[1], id: m[2] } : null;
}

function obtenerTokenSpotify() {
  const creds = Utilities.base64Encode(
    CONFIG.SPOTIFY_CLIENT_ID + ':' + CONFIG.SPOTIFY_CLIENT_SECRET
  );
  const resp = UrlFetchApp.fetch('https://accounts.spotify.com/api/token', {
    method: 'post',
    headers: { Authorization: `Basic ${creds}` },
    payload: 'grant_type=client_credentials',
    muteHttpExceptions: true,
  });
  return JSON.parse(resp.getContentText()).access_token;
}

function metricasSpotify(url) {
  const parsed = parsearSpotify(url);
  if (!parsed) return null;

  const token    = obtenerTokenSpotify();
  const endpoint = `https://api.spotify.com/v1/${parsed.tipo}s/${parsed.id}`;
  const resp     = UrlFetchApp.fetch(endpoint, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
  });
  const json = JSON.parse(resp.getContentText());

  // Tracks: popularity (0–100). Episodes: duration, description.
  // Plays reales solo disponibles en Spotify for Podcasters API.
  return {
    views:       parsed.tipo === 'track' ? json.popularity : null,
    reach:       null,
    likes:       null,
    comentarios: null,
    shares:      null,
    _nota:       parsed.tipo === 'track'
                   ? `Popularidad: ${json.popularity}/100`
                   : `Duración: ${Math.round((json.duration_ms || 0) / 60000)} min`,
  };
}


// ── CONFIGURAR ENCABEZADOS ──────────────────────────────────
function configurarEncabezados() {
  const hoja = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getActiveSheet();

  // Solo agrega encabezados de métricas en columnas I en adelante
  const encabezadosMetricas = [
    'Reproducciones', 'Alcance', 'Likes', 'Comentarios', 'Shares', 'Última actualización'
  ];

  const rango = hoja.getRange(1, CONFIG.COL_VIEWS, 1, encabezadosMetricas.length);
  rango.setValues([encabezadosMetricas]);
  rango.setFontWeight('bold');
  rango.setBackground('#1a1a2e');
  rango.setFontColor('#ffffff');
  hoja.setFrozenRows(1);

  Logger.log('✅ Encabezados de métricas agregados');
}

// ── INDEXADOR DE INSTAGRAM ──────────────────────────────────
// Corre una vez para llenar la columna IG Media ID en todas las filas
function indexarInstagram() {
  const ss   = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const cache = construirCacheInstagram();
  let encontrados = 0, omitidos = 0;

  ss.getSheets().forEach(hoja => {
    const datos = hoja.getDataRange().getValues();
    for (let i = CONFIG.FILA_INICIO - 1; i < datos.length; i++) {
      const link = (datos[i][CONFIG.COL_LINK - 1] || '').toString();
      if (!link.includes('instagram.com')) continue;

      // Si ya tiene ID guardado, saltar
      const existente = datos[i][CONFIG.COL_IG_MEDIA_ID - 1];
      if (existente) { omitidos++; continue; }

      const shortcode = parsearIdInstagram(link);
      if (!shortcode) continue;

      const resolved = cache[shortcode];
      if (resolved) {
        hoja.getRange(i + 1, CONFIG.COL_IG_MEDIA_ID).setValue(resolved.mediaId);
        encontrados++;
      } else {
        hoja.getRange(i + 1, CONFIG.COL_IG_MEDIA_ID).setValue('NOT_FOUND');
      }
      SpreadsheetApp.flush();
    }
  });

  Logger.log(`Indexación completa: ${encontrados} encontrados, ${omitidos} ya tenían ID`);
}


// ── FUNCIONES POR PESTAÑA ───────────────────────────────────
function procesarPestana(nombreHoja, filaInicio, filaFin) {
  const ss   = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) { Logger.log('Hoja no encontrada: ' + nombreHoja); return; }
  procesarHoja(hoja, filaInicio, filaFin);
}

// Procesa una pestaña completa en secuencia usando PropertiesService
// para guardar el progreso entre ejecuciones
function actualizarSIMAN() {
  const NOMBRE   = 'SIMAN';
  const TAMANO   = 100;
  const props    = PropertiesService.getScriptProperties();
  const ss       = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const hoja     = ss.getSheetByName(NOMBRE);
  if (!hoja) { Logger.log('Hoja no encontrada: ' + NOMBRE); return; }

  const totalFilas = hoja.getLastRow();
  let filaActual   = parseInt(props.getProperty('SIMAN_fila') || '2');

  if (filaActual > totalFilas) {
    Logger.log('✅ SIMAN completa. Reiniciando para próxima actualización.');
    props.setProperty('SIMAN_fila', '2');
    return;
  }

  const filaFin = Math.min(filaActual + TAMANO - 1, totalFilas);
  Logger.log(`Procesando SIMAN filas ${filaActual}–${filaFin} de ${totalFilas}`);

  procesarHoja(hoja, filaActual, filaFin);

  // Guardar progreso
  props.setProperty('SIMAN_fila', String(filaFin + 1));
  Logger.log(`Progreso guardado. Próxima ejecución desde fila ${filaFin + 1}`);
}

function reiniciarSIMAN() {
  PropertiesService.getScriptProperties().setProperty('SIMAN_fila', '2');
  Logger.log('🔄 Progreso SIMAN reiniciado desde fila 2');
}


// Diagnóstico: muestra qué shortcodes de Instagram no están en el caché
function diagnosticarInstagram() {
  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const cache = construirCacheInstagram();
  const noEncontrados = [];

  ss.getSheets().forEach(hoja => {
    const datos = hoja.getDataRange().getValues();
    for (let i = CONFIG.FILA_INICIO - 1; i < datos.length; i++) {
      const link = (datos[i][CONFIG.COL_LINK - 1] || '').toString();
      if (!link.includes('instagram.com')) continue;
      const shortcode = parsearIdInstagram(link);
      if (shortcode && !cache[shortcode]) {
        noEncontrados.push(`${hoja.getName()} fila ${i+1}: ${shortcode}`);
      }
    }
  });

  Logger.log(`Shortcodes no encontrados (${noEncontrados.length}):`);
  noEncontrados.forEach(s => Logger.log(s));
}


// ── FUNCIÓN DE PRUEBA ────────────────────────────────────────
function probarFacebook() {
  const url = 'https://www.facebook.com/reel/998180366486980';
  const resultado = metricasFacebook(url, 'reel');
  Logger.log(JSON.stringify(resultado));
}

function probarYouTube() {
  const url = 'https://www.youtube.com/live/gp9unooRbi4?si=90xz7CCDbfrg4J_C&t=1708';
  const resultado = metricasYouTube(url);
  Logger.log(JSON.stringify(resultado));
}

function probarInstagram() {
  const url = 'https://www.instagram.com/p/DYXSD5tR3Ox/';
  const resultado = metricasInstagram(url, 'post');
  Logger.log(JSON.stringify(resultado));
}
