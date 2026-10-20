/* Dimensionatore Rapido FV + Accumulo - porting web della logica del foglio DIMENSIONATORE
   (Dimensionatore_Rapido_FV_Accumulo_v7, copia BUILD-9D: Napoli NA, margine prudenziale rimosso).
   Ogni funzione riporta la cella Excel di riferimento. Nessun prezzo: il foglio PREZZI non e' validato (DEC-7).
   Uso: DIM.calcola(input, comuni) -> oggetto con i risultati; valori non calcolabili = null (n.d.). */
(function (root) {
  'use strict';

  // Parametri del metodo (DIMENSIONATORE C88:C95, Assunzioni C17:C21) - identici al file Excel
  var P = {
    OFFSET_FV: 0.95,        // C88
    AUTOSUFF_OBIETTIVO: 0.77, // C89
    TETTO_BATT_1000: 2.5,   // C90
    DIRETTO_BASE: 0.33,     // C91
    DIRETTO_PDC: 0.43,      // C92
    CAP_DIRETTO_PROD: 0.5,  // C93
    DOD: 0.9,               // C94
    GUADAGNO_MIN: 0.08,     // C95
    PROD_F1: 0.85, PROD_F2: 0.10, PROD_F3: 0.05, // Assunzioni C17:C19
    CICLI: 300, EFF_BATT: 0.75,                  // Assunzioni C20:C21
    TAGLIE: [5, 10, 15, 20],                     // F88:F91
    ADDON_PDC_KWP: 0,       // C42 (nel file e' la costante 0: add-on non applicato, DIM-05 aperto)
    ANNI: 25
  };

  // Ipotesi dichiarate (input con valore predefinito nel file Excel): il risultato le elenca.
  var IPOTESI_DEFAULT = {
    prezzo: 0.36,           // C14
    aumento: 0.03,          // C16
    f1: 0.30, f2: 0.35, f3: 0.35, // C19:C21
    kwhPdc: 7800,           // D25 (nota del file: "default 3.000" - divergenza DIM-04, decisione DEC-5)
    kwhEv: 2000,            // D26
    kwhInd: 500,            // D27
    kwhClima: 450           // D28 (per unita')
  };

  function norm(s) { return String(s == null ? '' : s).trim().toLowerCase(); }

  // C7:C9 e C39: ricerca esatta del comune (come VLOOKUP FALSE, senza distinzione maiuscole)
  function trovaComune(nome, comuni) {
    var q = norm(nome);
    if (!q) return null;
    for (var i = 0; i < comuni.length; i++) if (norm(comuni[i][0]) === q) return comuni[i];
    return null;
  }

  function ceilStep(x, step) { return Math.ceil(x / step - 1e-12) * step; }
  function num(x) { return (typeof x === 'number' && isFinite(x)) ? x : null; }
  function cumulato(annuo, i, anni) {   // serie geometrica ((1+i)^n - 1)/i; i = 0 -> n (limite, DIM-11)
    if (annuo == null) return null;
    if (!i) return annuo * anni;
    return annuo * (Math.pow(1 + i, anni) - 1) / i;
  }

  function calcola(inp, comuni) {
    var o = { ipotesi: [], avvisi: [], input: inp };
    var c = trovaComune(inp.comune, comuni);
    o.comune = c ? c[0] : null;
    o.provincia = c ? c[1] : null;           // C7
    o.regione = c ? c[2] : null;             // C8
    o.zona = c ? c[3] : null;                // C9
    if (!c) o.avvisi.push('Comune non trovato: producibilita\' di zona non determinabile.');

    // Prezzo energia (C14) e aumento annuo (C16): ipotesi se non indicati
    var prezzo = num(inp.prezzo);
    if (prezzo == null || prezzo <= 0) { prezzo = IPOTESI_DEFAULT.prezzo; o.ipotesi.push('prezzo energia ' + String(prezzo).replace('.', ',') + ' EUR/kWh (media ipotizzata, non da bolletta)'); }
    var aumento = num(inp.aumento);
    if (aumento == null) { aumento = IPOTESI_DEFAULT.aumento; o.ipotesi.push('aumento annuo prezzo energia 3% (ipotesi)'); }
    var f1 = num(inp.f1), f2 = num(inp.f2), f3 = num(inp.f3);
    if (f1 == null || f2 == null || f3 == null) {
      f1 = IPOTESI_DEFAULT.f1; f2 = IPOTESI_DEFAULT.f2; f3 = IPOTESI_DEFAULT.f3;
      o.ipotesi.push('ripartizione consumi per fasce 30/35/35 (ipotesi, fasce non indicate)');
    }
    o.sommaFasceOk = Math.round((f1 + f2 + f3) * 10000) / 10000 === 1;   // C22
    if (!o.sommaFasceOk) o.avvisi.push('La somma F1+F2+F3 deve essere 100%.');

    // C31 spesa annua ricostruita
    var spesa = num(inp.spesa);
    var tipo = inp.tipo || 'Annuale';
    o.spesaAnnua = spesa == null ? null : (tipo === 'Mensile' ? spesa * 12 : (tipo === 'Bimestrale' ? spesa * 6 : spesa));
    // C32 consumo base
    var reale = num(inp.consumoReale);
    if (reale != null && reale > 0) { o.consumoBase = reale; o.fonteConsumo = 'bolletta (kWh reali)'; }
    else if (o.spesaAnnua != null && o.spesaAnnua > 0) { o.consumoBase = o.spesaAnnua / prezzo; o.fonteConsumo = 'stimato dalla spesa / prezzo'; }
    else { o.consumoBase = null; o.fonteConsumo = 'non disponibile'; }

    // C34:C37 carichi futuri
    var kPdc = num(inp.kwhPdc), kEv = num(inp.kwhEv), kInd = num(inp.kwhInd), kCl = num(inp.kwhClima);
    if (inp.pdc && kPdc == null) { kPdc = IPOTESI_DEFAULT.kwhPdc; o.ipotesi.push('pompa di calore ' + kPdc + ' kWh/anno (valore del file)'); }
    if (inp.ev && kEv == null) { kEv = IPOTESI_DEFAULT.kwhEv; o.ipotesi.push('auto elettrica ' + kEv + ' kWh/anno (ipotesi)'); }
    if (inp.ind && kInd == null) { kInd = IPOTESI_DEFAULT.kwhInd; o.ipotesi.push('piano a induzione ' + kInd + ' kWh/anno (ipotesi)'); }
    var nClima = num(inp.clima) || 0;
    if (nClima > 0 && kCl == null) { kCl = IPOTESI_DEFAULT.kwhClima; o.ipotesi.push('climatizzatore ' + kCl + ' kWh/anno per unita\' (ipotesi)'); }
    o.extraPdc = inp.pdc ? kPdc : 0;        // C34
    o.extraEv = inp.ev ? kEv : 0;           // C35
    o.extraInd = inp.ind ? kInd : 0;        // C36
    o.extraClima = nClima * (kCl || 0);     // C37

    if (o.consumoBase == null) {
      o.avvisi.push('Inserire la spesa o il consumo annuo reale: senza consumi il dimensionamento non e\' calcolabile.');
      return o;
    }
    // C38
    o.consumoDim = o.consumoBase + o.extraPdc + o.extraEv + o.extraInd + o.extraClima;
    // C39 producibilita': valore dell'elenco comuni (nel file = media PVGIS della provincia), altrimenti zona (1100/1200/1300), altrimenti 1200
    if (c && num(c[4]) != null) { o.resa = c[4]; o.fonteResa = 'PVGIS media provinciale'; }
    else {
      o.resa = o.zona === 'Nord' ? 1100 : o.zona === 'Centro' ? 1200 : o.zona === 'Sud' ? 1300 : 1200;
      o.fonteResa = 'valore di zona (ipotesi)';
      o.ipotesi.push('producibilita\' ' + o.resa + ' kWh/kWp di zona (PVGIS non disponibile)');
    }
    // C40:C44
    o.fvTeorico = o.consumoDim * P.OFFSET_FV / o.resa;
    o.fvArrotondato = ceilStep(o.fvTeorico, 0.5);
    var lordo = o.fvArrotondato + P.ADDON_PDC_KWP;
    o.fv = Math.min(10, Math.max(3, lordo));
    o.avvisoTaglia = lordo > 10 ? 'Fabbisogno > 10 kWp: progettazione dedicata fuori range residenziale standard'
      : (lordo < 3 ? 'Fabbisogno < 3 kWp: applicata taglia minima 3 kWp' : 'Dimensionamento nel range residenziale 3-10 kWp');
    // C46:C51
    o.kwhF1 = o.consumoDim * f1; o.kwhF2 = o.consumoDim * f2; o.kwhF3 = o.consumoDim * f3;
    o.kwhF23 = o.kwhF2 + o.kwhF3;
    o.accumuloTeorico = (o.kwhF23 * 0.8) / 365;
    o.accumuloMinPdc = inp.pdc ? 5 : 0;      // C51 (informativo)
    // C55
    o.produzione = o.fv * o.resa;
    // C96:C101
    o.dirApplicato = inp.pdc ? P.DIRETTO_PDC : P.DIRETTO_BASE;
    o.autoDiretto = Math.min(o.consumoDim * o.dirApplicato, o.produzione * P.CAP_DIRETTO_PROD);
    o.consumoSerale = o.consumoDim * (f2 + f3);
    o.surplus = Math.max(0, o.produzione - o.autoDiretto);
    o.caricabile = Math.min(o.consumoSerale, o.surplus);
    o.tettoBatteria = P.TETTO_BATT_1000 * o.consumoDim / 1000;
    // G88:G91 e C52
    function autosuff(k) { return Math.min(1, (o.autoDiretto + o.caricabile * (1 - Math.exp(-(k * P.DOD) / (o.consumoSerale / 365)))) / o.consumoDim); }
    var T = P.TAGLIE, G = T.map(autosuff);
    o.autosuffTaglie = G;
    var acc;
    if (G[0] >= P.AUTOSUFF_OBIETTIVO || T[1] > o.tettoBatteria || (G[1] - G[0]) < P.GUADAGNO_MIN) acc = T[0];
    else if (G[1] >= P.AUTOSUFF_OBIETTIVO || T[2] > o.tettoBatteria || (G[2] - G[1]) < P.GUADAGNO_MIN) acc = T[1];
    else if (G[2] >= P.AUTOSUFF_OBIETTIVO || T[3] > o.tettoBatteria || (G[3] - G[2]) < P.GUADAGNO_MIN) acc = T[2];
    else acc = T[3];
    o.accumulo = acc;
    // C63:C72
    o.prodF1 = o.produzione * P.PROD_F1; o.prodF2 = o.produzione * P.PROD_F2; o.prodF3 = o.produzione * P.PROD_F3;
    o.dirF1 = Math.min(o.prodF1, o.kwhF1); o.dirF2 = Math.min(o.prodF2, o.kwhF2); o.dirF3 = Math.min(o.prodF3, o.kwhF3);
    o.surF1 = Math.max(0, o.prodF1 - o.dirF1); o.surF2 = Math.max(0, o.prodF2 - o.dirF2); o.surF3 = Math.max(0, o.prodF3 - o.dirF3);
    o.risparmioSenzaKwh = o.autoDiretto;                                  // C67
    o.capacitaBatteria = acc * P.CICLI;                                   // C69
    o.spostataBatteria = Math.min(o.surF1 * P.EFF_BATT, (o.kwhF2 - o.dirF2) + (o.kwhF3 - o.dirF3), acc * P.CICLI * P.EFF_BATT); // C70 (informativo)
    o.risparmioConKwh = Math.min(o.consumoDim, o.autoDiretto + o.caricabile * (1 - Math.exp(-(acc * P.DOD) / (o.consumoSerale / 365)))); // C71
    o.coperturaSenza = o.risparmioSenzaKwh / o.consumoDim;                // C72
    o.coperturaCon = o.risparmioConKwh / o.consumoDim;                    // D72
    // C75:C85
    o.bollettaAnnua = (reale != null && reale > 0) ? reale * prezzo : o.spesaAnnua;
    if (o.bollettaAnnua == null) o.avvisi.push('Bolletta non calcolabile: manca la spesa.');
    o.bollettaMese = o.bollettaAnnua == null ? null : o.bollettaAnnua / 12;
    o.bollettaGiorno = o.bollettaAnnua == null ? null : o.bollettaAnnua / 365;
    o.risparmioSenzaEur = o.risparmioSenzaKwh * prezzo;                   // C79
    o.risparmioConEur = o.risparmioConKwh * prezzo;                       // C80
    o.costoGiornoSenza = o.risparmioSenzaEur / 365;                       // C83
    o.costo25Senza = cumulato(o.risparmioSenzaEur, aumento, P.ANNI);      // D83
    o.costoGiornoCon = o.risparmioConEur / 365;                           // C84
    o.costo25Con = cumulato(o.risparmioConEur, aumento, P.ANNI);          // D84
    o.bolletta25 = cumulato(o.bollettaAnnua, aumento, P.ANNI);            // C85
    o.prezzoKit = null;   // C57/C58 non portati: PREZZI senza fonte validata (DEC-7)
    return o;
  }

  // FIXED() di Excel in locale italiano: arrotondamento "per eccesso a meta'", punto per le migliaia, virgola per i decimali.
  function fixedIt(x, dec, senzaMigliaia) {
    if (x == null || !isFinite(x)) return null;
    var f = Math.pow(10, dec), r = Math.sign(x) * Math.round(Math.abs(x) * f + 1e-9) / f;
    var s = Math.abs(r).toFixed(dec), p = s.split('.'), i = p[0];
    if (!senzaMigliaia) i = i.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return (r < 0 ? '-' : '') + i + (dec > 0 ? ',' + p[1] : '');
  }

  // SINTESI_CLIENTE!B31 (copia RELEASE): stessa frase, stessi arrotondamenti; senza consumi -> frase non disponibile.
  function frase(o) {
    if (!o || o.consumoDim == null || !(o.consumoDim > 0)) return 'Frase non disponibile: mancano i dati di consumo o di spesa del cliente.';
    var aum = o.input && o.input.aumento != null ? o.input.aumento : IPOTESI_DEFAULT.aumento;
    var v = [fixedIt(o.bollettaGiorno, 2), fixedIt(aum * 100, 0, true), fixedIt(o.bolletta25, 0), fixedIt(o.fv, 1), fixedIt(o.accumulo, 0, true),
      fixedIt(o.produzione, 0), fixedIt(o.coperturaCon * 100, 0, true), fixedIt(o.costoGiornoSenza, 2), fixedIt(o.costo25Senza, 0),
      fixedIt(o.costoGiornoCon, 2), fixedIt(o.costo25Con, 0)];
    if (v.some(function (x) { return x == null; })) return 'Frase non disponibile: verificare i dati inseriti.';
    return 'In base ai dati disponibili, il cliente sostiene una bolletta di circa € ' + v[0] + ' al giorno. Considerando un aumento del ' + v[1] +
      '% annuo del costo dell\'energia, nei prossimi 25 anni la bolletta cumulata supererà i € ' + v[2] + '. L\'impianto consigliato (' + v[3] +
      ' kWp + ' + v[4] + ' kWh) produce circa ' + v[5] + ' kWh/anno e copre il ' + v[6] + '% del fabbisogno con accumulo. Il vero costo del rinvio è il risparmio che non viene attivato: € ' +
      v[7] + '/giorno con il solo FV (€ ' + v[8] + ' su 25 anni) e fino a € ' + v[9] + '/giorno con FV + accumulo (€ ' + v[10] + ' su 25 anni).';
  }

  var api = { calcola: calcola, frase: frase, fixedIt: fixedIt, trovaComune: trovaComune, PARAMETRI: P, IPOTESI_DEFAULT: IPOTESI_DEFAULT, VERSIONE: '2026.10.18-REL' };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.DIM = api;
})(this);
