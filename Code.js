const CONFIG = {
  SPREADSHEET_ID: '1aVTYW8cXcXxFGIatcpSpBlpods2pV_mrRnYo-CEV2m8',
  VINYL_SHEET: 'Vinilos',
  RECOMMENDATIONS_SHEET: 'Recomendaciones Web',
  RECENT_RECOMMENDATION_DAYS: 30,
  CACHE_SECONDS: 300,
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Colección de Vinilos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getBootstrapData() {
  // No cacheamos el payload completo: CacheService limita cada valor a ~100 KB
  // y la colección completa supera ese tamaño. Las portadas de Discogs sí
  // siguen usando caché individual en resolveCoverUrl_().
  const albums = getCollection_();

  return {
    generatedAt: new Date().toISOString(),
    collection: albums,
    stats: computeStats_(albums),
    recommendations: getOrCreateTodaysRecommendations_(albums),
  };
}

function refreshBootstrapData() {
  return getBootstrapData();
}

function markRecommendationListened(albumId) {
  const sheet = ensureRecommendationsSheet_();
  const rows = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  for (let r = rows.length - 1; r >= 1; r--) {
    if (String(rows[r][0]) === today && String(rows[r][1]) === String(albumId)) {
      sheet.getRange(r + 1, 5).setValue('Escuchado');
      sheet.getRange(r + 1, 6).setValue(new Date());
      break;
    }
  }
  return { ok: true };
}

function replaceRecommendation(albumId) {
  const albums = getCollection_();
  const sheet = ensureRecommendationsSheet_();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const rows = sheet.getDataRange().getValues();
  const excluded = new Set([String(albumId)]);

  for (let r = 1; r < rows.length; r++) {
    if (String(rows[r][0]) === today) excluded.add(String(rows[r][1]));
  }

  const history = getRecommendationHistory_();
  const candidate = weightedPick_(albums, history, excluded);
  if (!candidate) return null;

  for (let r = rows.length - 1; r >= 1; r--) {
    if (String(rows[r][0]) === today && String(rows[r][1]) === String(albumId)) {
      sheet.getRange(r + 1, 5).setValue('Reemplazado');
      sheet.getRange(r + 1, 6).setValue(new Date());
      break;
    }
  }

  appendRecommendation_(candidate, 'Alternativa');
  CacheService.getScriptCache().remove('bootstrap-v1');
  return enrichAlbum_(candidate);
}

function getCollection_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.VINYL_SHEET);
  if (!sheet) throw new Error(`No existe la hoja ${CONFIG.VINYL_SHEET}`);

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h).trim());
  const col = {};
  headers.forEach((h, i) => col[h] = i);

  const field = (row, name) => col[name] === undefined ? '' : row[col[name]];
  const num = value => {
    if (value === '' || value === null || value === undefined) return null;
    const n = Number(value);
    return Number.isFinite(n) ? n : null;
  };

  return values.slice(1)
    .filter(row => field(row, 'Título') && field(row, 'Artista'))
    .filter(row => String(field(row, 'Estado') || '').trim().toLowerCase() !== 'canjeado')
    .map(row => {
      const purchaseDate = field(row, 'Fecha Compra');
      return {
        id: String(field(row, 'Id') || ''),
        purchaseDate: normalizeDate_(purchaseDate),
        status: String(field(row, 'Estado') || ''),
        title: String(field(row, 'Título') || '').trim(),
        artist: String(field(row, 'Artista') || '').trim(),
        collection: String(field(row, 'Coleccion') || '').trim(),
        origin: String(field(row, 'Origen') || '').trim(),
        purchasedAt: String(field(row, 'CompradoEn') || '').trim(),
        releaseYear: num(field(row, 'Año Lanzamiento')),
        editionYear: num(field(row, 'Año Edición')),
        features: String(field(row, 'Características') || '').trim(),
        discogsUrl: String(field(row, 'DiscoGS') || '').trim(),
        condition: String(field(row, 'Condicion') || '').trim(),
        personalMemory: num(field(row, 'Memoria personal')),
        sharedEnjoyment: num(field(row, 'Disfrute compartido')),
        curatorialValue: num(field(row, 'Valor curatorial')),
        editionInterest: num(field(row, 'Interés de edición')),
        copyQuality: num(field(row, 'Calidad de copia')),
        listeningProbability: num(field(row, 'Probabilidad de escucha')),
        note: String(field(row, 'Nota curatorial') || '').trim(),
      };
    })
    .map(enrichAlbum_);
}

function enrichAlbum_(album) {
  return Object.assign({}, album, {
    coverUrl: resolveCoverUrl_(album),
    displayArtist: normalizeArtistDisplay_(album.artist),
  });
}

function resolveCoverUrl_(album) {
  if (!album.discogsUrl) return '';
  const match = album.discogsUrl.match(/\/release\/(\d+)/i);
  if (!match) return '';

  const token = PropertiesService.getScriptProperties().getProperty('DISCOGS_TOKEN');
  if (!token) return '';

  const cache = CacheService.getScriptCache();
  const key = `cover-${match[1]}`;
  const cached = cache.get(key);
  if (cached) return cached === '__none__' ? '' : cached;

  try {
    const response = UrlFetchApp.fetch(`https://api.discogs.com/releases/${match[1]}`, {
      muteHttpExceptions: true,
      headers: {
        Authorization: `Discogs token=${token}`,
        'User-Agent': 'ColeccionVinilosSites/1.0',
      },
    });
    if (response.getResponseCode() !== 200) {
      cache.put(key, '__none__', 21600);
      return '';
    }
    const json = JSON.parse(response.getContentText());
    const image = (json.images || []).find(x => x.type === 'primary') || (json.images || [])[0];
    const url = image ? image.uri500 || image.uri || '' : '';
    cache.put(key, url || '__none__', 21600);
    return url;
  } catch (e) {
    return '';
  }
}

function computeStats_(albums) {
  const byArtist = new Map();
  const byDecade = {};
  const byOrigin = {};
  const byStore = {};
  const purchasesByYear = {};

  albums.forEach(a => {
    const key = normalizeArtistKey_(a.artist);
    if (!byArtist.has(key)) {
      byArtist.set(key, {
        artist: normalizeArtistDisplay_(a.artist),
        copies: 0,
        titles: new Set(),
      });
    }
    const rec = byArtist.get(key);
    rec.copies += 1;
    rec.titles.add(normalizeTitleKey_(a.title));

    if (a.releaseYear) {
      const decade = `${Math.floor(a.releaseYear / 10) * 10}s`;
      byDecade[decade] = (byDecade[decade] || 0) + 1;
    }

    const origin = a.origin || 'Sin dato';
    byOrigin[origin] = (byOrigin[origin] || 0) + 1;

    const store = a.purchasedAt || 'Sin dato';
    byStore[store] = (byStore[store] || 0) + 1;

    if (a.purchaseDate) {
      const year = a.purchaseDate.slice(0, 4);
      purchasesByYear[year] = (purchasesByYear[year] || 0) + 1;
    }
  });

  const artists = [...byArtist.values()].map(x => ({
    artist: x.artist,
    copies: x.copies,
    uniqueAlbums: x.titles.size,
  })).sort((a, b) => b.copies - a.copies || a.artist.localeCompare(b.artist));

  const uniqueAlbumCount = [...byArtist.values()].reduce((sum, x) => sum + x.titles.size, 0);
  const artistCount = byArtist.size;

  return {
    totalCopies: albums.length,
    uniqueAlbums: uniqueAlbumCount,
    artistCount,
    avgCopiesPerArtist: artistCount ? round_(albums.length / artistCount, 2) : 0,
    avgUniqueAlbumsPerArtist: artistCount ? round_(uniqueAlbumCount / artistCount, 2) : 0,
    artists,
    byDecade: objectToSeries_(byDecade),
    byOrigin: objectToSeries_(byOrigin),
    byStore: objectToSeries_(byStore).sort((a,b) => b.value - a.value),
    purchasesByYear: objectToSeries_(purchasesByYear).sort((a,b) => a.label.localeCompare(b.label)),
  };
}

function getOrCreateTodaysRecommendations_(albums) {
  const sheet = ensureRecommendationsSheet_();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const rows = sheet.getDataRange().getValues();
  const todayIds = rows.slice(1)
    .filter(r => String(r[0]) === today && ['Pendiente', 'Escuchado'].includes(String(r[4] || 'Pendiente')))
    .map(r => String(r[1]));

  if (todayIds.length) {
    return todayIds.map(id => albums.find(a => String(a.id) === id)).filter(Boolean);
  }

  const history = getRecommendationHistory_();
  const daysSinceLast = daysSinceLastRecommendation_(history);
  const total = 1 + extraRecommendationCount_(daysSinceLast);
  const selected = [];
  const excluded = new Set();

  for (let i = 0; i < total; i++) {
    const pick = weightedPick_(albums, history, excluded);
    if (!pick) break;
    selected.push(pick);
    excluded.add(String(pick.id));
    appendRecommendation_(pick, i === 0 ? 'Del día' : 'Adicional');
  }
  return selected;
}

function getRecommendationHistory_() {
  const sheet = ensureRecommendationsSheet_();
  const rows = sheet.getDataRange().getValues();
  return rows.slice(1).map(r => ({
    date: String(r[0]),
    albumId: String(r[1]),
    type: String(r[3]),
    status: String(r[4]),
  }));
}

function ensureRecommendationsSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.RECOMMENDATIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.RECOMMENDATIONS_SHEET);
    sheet.getRange(1, 1, 1, 6).setValues([[
      'Fecha', 'Id Vinilo', 'Título / Artista', 'Tipo', 'Estado', 'Actualizado'
    ]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function appendRecommendation_(album, type) {
  const sheet = ensureRecommendationsSheet_();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sheet.appendRow([
    today,
    album.id,
    `${album.title} — ${normalizeArtistDisplay_(album.artist)}`,
    type,
    'Pendiente',
    new Date(),
  ]);
}

function weightedPick_(albums, history, excluded) {
  const now = new Date();
  const recentCutoff = new Date(now.getTime() - CONFIG.RECENT_RECOMMENDATION_DAYS * 86400000);
  const recentIds = new Set(history
    .filter(h => parseHistoryDate_(h.date) >= recentCutoff)
    .map(h => String(h.albumId)));

  let candidates = albums.filter(a => a.id && !excluded.has(String(a.id)) && !recentIds.has(String(a.id)));
  if (!candidates.length) candidates = albums.filter(a => a.id && !excluded.has(String(a.id)));
  if (!candidates.length) return null;

  const weighted = candidates.map(a => {
    const memory = a.personalMemory || 2.5;
    const curatorial = a.curatorialValue || 3;
    const listen = a.listeningProbability || 3;
    const edition = a.editionInterest || 2;
    const copyQuality = a.copyQuality || 3;
    const noteBonus = a.note ? 0.4 : 0;
    const ageBonus = recommendationAgeBonus_(a.id, history);
    const qualityPenalty = copyQuality <= 1 ? 0.5 : 1;
    const score = Math.max(0.2, (memory * 0.28 + curatorial * 0.32 + listen * 0.28 + edition * 0.12 + noteBonus + ageBonus) * qualityPenalty);
    return { album: a, weight: score * score };
  });

  const total = weighted.reduce((s, x) => s + x.weight, 0);
  let target = Math.random() * total;
  for (const item of weighted) {
    target -= item.weight;
    if (target <= 0) return item.album;
  }
  return weighted[weighted.length - 1].album;
}

function recommendationAgeBonus_(albumId, history) {
  const dates = history
    .filter(h => String(h.albumId) === String(albumId))
    .map(h => parseHistoryDate_(h.date))
    .filter(Boolean)
    .sort((a,b) => b-a);
  if (!dates.length) return 1.2;
  const days = Math.floor((new Date() - dates[0]) / 86400000);
  return Math.min(1.5, days / 120);
}

function daysSinceLastRecommendation_(history) {
  const dates = history.map(h => parseHistoryDate_(h.date)).filter(Boolean).sort((a,b) => b-a);
  if (!dates.length) return 999;
  return Math.max(0, Math.floor((new Date() - dates[0]) / 86400000));
}

function extraRecommendationCount_(days) {
  if (days <= 2) return 0;
  if (days <= 6) return 1;
  if (days <= 13) return 2;
  return 3;
}

function normalizeDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const raw = String(value).trim();
  const m = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  return raw;
}

function parseHistoryDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]') return value;
  const s = String(value);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

function normalizeArtistDisplay_(artist) {
  const s = String(artist || '').trim();
  if (!s) return '';
  if (s.endsWith(' - The')) return `The ${s.slice(0, -6)}`;
  const person = s.match(/^(.+?)\s-\s(.+)$/);
  if (person && !s.includes(',')) return `${person[2]} ${person[1]}`;
  return s;
}

function normalizeArtistKey_(artist) {
  return normalizeArtistDisplay_(artist).toLowerCase().replace(/[^a-z0-9áéíóúüñ]+/gi, ' ').trim();
}

function normalizeTitleKey_(title) {
  return String(title || '').toLowerCase().replace(/\([^)]*\)/g, '').replace(/[^a-z0-9áéíóúüñ]+/gi, ' ').trim();
}

function objectToSeries_(obj) {
  return Object.entries(obj).map(([label, value]) => ({ label, value }));
}

function round_(n, digits) {
  const p = Math.pow(10, digits);
  return Math.round(n * p) / p;
}
