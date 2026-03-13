// =============================================
// bp-core.js — Shared data logic for HomeBP Insight
// Used by both clinician (index.html) and patient (patient.html) versions.
// This module is DOM-free: all UI config values are passed as parameters.
// =============================================

const BPCore = (() => {
  'use strict';

  // =============================================
  // CONSTANTS & CONFIG
  // =============================================
  const COLOR_GREEN = '#00C853';
  const COLOR_RED = '#FF4444';
  const COLOR_SEVERE = '#FF0000';
  const COLOR_HYPO = '#7B1FA2';
  const HR_COLOR = '#8a8a8a';
  const TARGET_COLOR = 'rgba(0,0,0,.25)';
  const HOLD_STAR_COLOR = '#000000';

  const TARGET_SBP = 140, TARGET_DBP = 90;
  const GREEN_SBP_MAX = 140, GREEN_DBP_MAX = 90;
  const SEVERE_SBP_MIN = 160, SEVERE_DBP_MIN = 100;
  const SEVERE_LABEL_SBP_MIN = 160;
  const HYPO_SBP_MAX = 100, HYPO_DBP_MAX = 60;

  const BASE_BAR_W = 0.28;
  const SEVERE_BAR_W = 0.44;

  const MED_PALETTE = ['#2f6db3','#1b9e77','#d95f02','#7570b3','#66a61e','#e7298a','#a6761d','#e6ab02','#1f78b4','#b2df8a','#fb9a99','#cab2d6'];

  const ALPHA_MIN = 0.15, ALPHA_MAX = 0.85;
  const DOSE_RATIO_MAX = 4;

  const STARTING_DOSES = { valsartan:80, candesartan:8, irbesartan:150, metoprolol:12.5 };

  const FREQ_CANON = {
    'qd':'qd','daily':'qd','once daily':'qd',
    'bid':'bid','twice daily':'bid',
    'tid':'tid','qid':'qid','qhs':'qhs','qod':'qod',
    'q d':'qd','b i d':'bid','t i d':'tid',
  };

  const CANVAS_BG = [1, 1, 1];

  // =============================================
  // UTILITY FUNCTIONS
  // =============================================
  function mean(nums) { if (!nums.length) return null; let s=0; for (const n of nums) s+=n; return s/nums.length; }
  function stddev(nums) { if (nums.length<2) return null; const m=mean(nums); let v=0; for (const n of nums) v+=(n-m)**2; return Math.sqrt(v/(nums.length-1)); }
  function quantile(arr, q) {
    if (!arr.length) return null;
    const sorted = [...arr].sort((a,b)=>a-b);
    const pos = (sorted.length-1)*q;
    const base = Math.floor(pos);
    const rest = pos - base;
    return sorted[base+1]!==undefined ? sorted[base]+rest*(sorted[base+1]-sorted[base]) : sorted[base];
  }
  function fmtBP(sys,dia) { return (sys==null||dia==null)?'—':`${Math.round(sys)}/${Math.round(dia)}`; }
  function fmtDate(d) { return d.toLocaleString(undefined,{year:'numeric',month:'short',day:'2-digit',hour:'2-digit',minute:'2-digit'}); }
  function fmtDateShort(d) { const mo=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']; return `${mo[d.getMonth()]} ${d.getDate()}`; }
  function fmtPct(x) { return x==null||isNaN(x)?'—':`${x.toFixed(1)}%`; }
  function fmtNum(x) { return x==null||isNaN(x)?'—':x.toFixed(1); }
  function daysBetween(a,b) { return (b.getTime()-a.getTime())/(86400000); }

  // =============================================
  // CSV PARSER
  // =============================================
  function parseCSV(text) {
    const rows=[]; let i=0,field='',row=[],inQ=false;
    const pf=()=>{row.push(field);field='';};
    const pr=()=>{rows.push(row);row=[];};
    while(i<text.length){const c=text[i];if(inQ){if(c==='"'){if(text[i+1]==='"'){field+='"';i+=2;continue;}inQ=false;i++;continue;}field+=c;i++;continue;}if(c==='"'){inQ=true;i++;continue;}if(c===','){pf();i++;continue;}if(c==='\n'){pf();pr();i++;continue;}if(c==='\r'){i++;continue;}field+=c;i++;}
    pf();if(row.length>1||row[0]!=='')pr();return rows;
  }

  // =============================================
  // EXCEL PARSER
  // =============================================
  function normalizeHeader(h) { return String(h||'').trim().toLowerCase().replace(/[^a-z0-9]/g,''); }
  function sheetToAoa(wb) { const s=wb.Sheets[wb.SheetNames[0]]; return XLSX.utils.sheet_to_json(s,{header:1,defval:''}); }
  function findLikelyHeaderRow(aoa) {
    const cands=aoa.slice(0,40);
    const score=row=>{const cells=row.map(normalizeHeader);let s=0;const has=arr=>arr.some(x=>cells.includes(normalizeHeader(x)));
      if(has(['datetime','datetimelocal','date/time','timestamp','date']))s+=2;
      if(has(['systolic','sys','sbp','systolic(mmhg)','systolicmmhg']))s+=2;
      if(has(['diastolic','dia','dbp','diastolic(mmhg)','diastolicmmhg']))s+=2;
      if(has(['pulse','hr','heartrate','pulsebpm','pulse(bpm)']))s+=1;
      if(row.filter(v=>String(v).trim()!=='').length>=3)s+=1;return s;};
    let bi=0,bs=-1;for(let i=0;i<cands.length;i++){const sc=score(cands[i]||[]);if(sc>bs){bs=sc;bi=i;}}
    return bs>=5?bi:0;
  }
  function excelAoaToRows(aoa) {
    const hi=findLikelyHeaderRow(aoa);const hdr=(aoa[hi]||[]).map(h=>String(h).trim());
    const data=aoa.slice(hi+1).filter(r=>r.some(v=>String(v).trim()!==''));
    const rows=[hdr];for(const r of data){const rr=r.slice();while(rr.length<hdr.length)rr.push('');rows.push(rr);}return rows;
  }

  // =============================================
  // DATE PARSING
  // =============================================
  function tryParseDate(s) {
    if(s==null)return null;
    if(Object.prototype.toString.call(s)==='[object Date]'&&!isNaN(s.getTime()))return s;
    const raw=String(s).trim();if(!raw)return null;
    if(typeof s==='number'&&isFinite(s)){const d=XLSX.SSF.parse_date_code(s);if(d&&d.y)return new Date(d.y,(d.m||1)-1,d.d||1,d.H||0,d.M||0,d.S||0);}
    if(/^\d{10,13}$/.test(raw)){const n=Number(raw);const ms=raw.length===10?n*1000:n;const d=new Date(ms);return isNaN(d.getTime())?null:d;}
    const d1=new Date(raw);if(!isNaN(d1.getTime()))return d1;
    const m=raw.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);
    if(m)return new Date(+m[1],+m[2]-1,+m[3],+m[4],+m[5],+(m[6]||0));
    return null;
  }

  function parseDateAndTime(dateVal, timeVal) {
    // Handle Excel serial date numbers (e.g. 46083.31162 → 2026-03-02 07:28)
    if (typeof dateVal === 'number' && isFinite(dateVal) && dateVal > 1) {
      if (typeof XLSX !== 'undefined' && XLSX.SSF && XLSX.SSF.parse_date_code) {
        const d = XLSX.SSF.parse_date_code(dateVal);
        if (d && d.y) return new Date(d.y, (d.m||1)-1, d.d||1, d.H||0, d.M||0, d.S||0);
      }
      const jsDate = new Date(Math.round((dateVal - 25569) * 86400000));
      if (!isNaN(jsDate.getTime())) return jsDate;
    }
    const dStr = String(dateVal||'').trim();
    const tStr = String(timeVal||'').trim();
    if (!dStr) return null;
    const combined = tStr ? `${dStr} ${tStr}` : dStr;
    return tryParseDate(combined);
  }

  // =============================================
  // COLUMN DETECTION
  // =============================================
  function detectColumns(headers) {
    const norms = headers.map(h=>normalizeHeader(h));
    const findAny = cands => { for(const c of cands){const i=norms.indexOf(normalizeHeader(c));if(i!==-1)return i;} return -1; };

    const dtIdx = findAny(['DateTimeLocal','DateTime','Timestamp','Date/Time','Datetime']);
    const dateIdx = findAny(['Date']);
    const timeIdx = findAny(['Time']);
    const sysIdx = findAny(['Systolic','SYS','SBP','Systolic(mmHg)','Systolicmmhg','SBP(mmHg)','SBPmmhg']);
    const diaIdx = findAny(['Diastolic','DIA','DBP','Diastolic(mmHg)','Diastolicmmhg','DBP(mmHg)','DBPmmhg']);
    const pulseIdx = findAny(['Pulse','HR','Heartrate','HeartRate','Pulse(bpm)','Pulsebpm','Pulse(bpm)']);
    const notesIdx = findAny(['Notes','Note','Comment','Comments','Annotation','Annotations']);

    const hasDT = dtIdx !== -1;
    const hasDatePlusTime = dateIdx !== -1;
    const hasBP = sysIdx !== -1 && diaIdx !== -1;

    const ok = (hasDT || hasDatePlusTime) && hasBP;

    return {
      ok,
      useCombinedDT: hasDT,
      dtIdx: hasDT ? dtIdx : -1,
      dateIdx,
      timeIdx,
      sysIdx,
      diaIdx,
      pulseIdx: pulseIdx !== -1 ? pulseIdx : null,
      notesIdx: notesIdx !== -1 ? notesIdx : null,
    };
  }

  // =============================================
  // BP CLASSIFICATION
  // =============================================
  function classifyBP(sys, dia) {
    const hypo = sys < HYPO_SBP_MAX || dia < HYPO_DBP_MAX;
    const severe = sys >= SEVERE_SBP_MIN || dia >= SEVERE_DBP_MIN;
    const green = !hypo && sys < GREEN_SBP_MAX && dia < GREEN_DBP_MAX && !severe;
    const red = !green && !severe && !hypo;
    return { severe, green, red, hypo };
  }

  // Re-tag .green and .red on readings based on a custom goal (e.g. US 130/80 or custom).
  // Hypo and severe are independent of goal and remain unchanged.
  function reclassifyForGoal(readings, goal) {
    for (const r of readings) {
      r.green = !r.hypo && r.sys < goal.sys && r.dia < goal.dia && !r.severe;
      r.red = !r.green && !r.severe && !r.hypo;
    }
  }

  // =============================================
  // BUILD READINGS
  // Accepts config: { amStart, amEnd } — passed from UI
  // =============================================
  function buildReadings(rawRows, mapping, config) {
    const amStart = config.amStart || 0;
    const amEnd = config.amEnd || 12;
    let result = [];

    if (!rawRows || rawRows.length < 2 || !mapping) return result;

    for (let r = 1; r < rawRows.length; r++) {
      const row = rawRows[r];
      let t;
      if (mapping.useCombinedDT && mapping.dtIdx >= 0) {
        t = tryParseDate(row[mapping.dtIdx]);
      } else {
        t = parseDateAndTime(row[mapping.dateIdx], mapping.timeIdx >= 0 ? row[mapping.timeIdx] : '');
      }
      const sys = Number(String(row[mapping.sysIdx]??'').trim());
      const dia = Number(String(row[mapping.diaIdx]??'').trim());
      const pulse = mapping.pulseIdx!=null ? Number(String(row[mapping.pulseIdx]??'').trim()) : null;
      const notes = mapping.notesIdx!=null ? String(row[mapping.notesIdx]||'').trim() : '';

      if (!t || !Number.isFinite(sys) || !Number.isFinite(dia)) continue;
      if (sys < 60 || sys > 260 || dia < 30 || dia > 160) continue;

      const hour = t.getHours();
      let window;
      if (hour >= amStart && hour < amEnd) window = 'Morning';
      else if (hour >= 12 && hour < 24) window = 'Afternoon/Evening';
      else window = 'Other';

      const cls = classifyBP(sys, dia);
      result.push({
        t, sys, dia,
        hr: Number.isFinite(pulse) ? pulse : null,
        notes,
        window,
        green: cls.green, red: cls.red, severe: cls.severe, hypo: cls.hypo,
        dateOnly: new Date(t.getFullYear(), t.getMonth(), t.getDate()),
      });
    }
    result.sort((a,b) => a.t - b.t);

    // Deduplicate exact duplicates
    const deduped = []; let prevKey = null;
    for (const x of result) {
      const key = `${x.t.getTime()}|${x.sys}|${x.dia}|${x.hr??''}`;
      if (prevKey === key) continue;
      deduped.push(x); prevKey = key;
    }

    // Merge readings within 10 minutes
    const MERGE_WINDOW_MS = 10 * 60 * 1000;
    const merged = [];
    let i = 0;
    while (i < deduped.length) {
      let j = i + 1;
      while (j < deduped.length && (deduped[j].t.getTime() - deduped[i].t.getTime()) <= MERGE_WINDOW_MS) {
        j++;
      }
      if (j - i === 1) {
        merged.push(deduped[i]);
      } else {
        const group = deduped.slice(i, j);
        const avgSys = Math.round(mean(group.map(r => r.sys)));
        const avgDia = Math.round(mean(group.map(r => r.dia)));
        const hrs = group.filter(r => r.hr != null).map(r => r.hr);
        const avgHR = hrs.length ? Math.round(mean(hrs)) : null;
        const midTime = new Date(Math.round(mean(group.map(r => r.t.getTime()))));
        const notes = group.map(r => r.notes).filter(Boolean).join('; ');
        const cls = classifyBP(avgSys, avgDia);
        const hour = midTime.getHours();
        let win;
        if (hour >= amStart && hour < amEnd) win = 'Morning';
        else if (hour >= 12 && hour < 24) win = 'Afternoon/Evening';
        else win = 'Other';
        merged.push({
          t: midTime, sys: avgSys, dia: avgDia, hr: avgHR, notes,
          window: win,
          green: cls.green, red: cls.red, severe: cls.severe, hypo: cls.hypo,
          dateOnly: new Date(midTime.getFullYear(), midTime.getMonth(), midTime.getDate()),
          mergedCount: group.length,
          components: group.map(r => ({ t: r.t, sys: r.sys, dia: r.dia, hr: r.hr })),
        });
      }
      i = j;
    }
    return merged;
  }

  function findHoldEventsFromReadings(readings) {
    const holdEvents = [];
    for (const r of readings) {
      const n = r.notes.toLowerCase();
      if (n.includes('hold') && (n.includes('med') || n.includes('htn') || n.includes('bp'))) {
        holdEvents.push(r.t);
      }
    }
    return holdEvents;
  }

  // =============================================
  // MEDICATION PARSING ENGINE
  // =============================================
  function splitNoteLines(note) {
    return note.split(/[\n\r]+|(?:\s*\/\s*(?=[A-Z]))|(?:\s*;\s*)/)
      .map(s => s.replace(/^\s*(?:meds\s*:|comment\s*:)\s*/i, '').trim())
      .filter(s => s.length > 0);
  }

  const MED_START_RE = /\b(start|started|begin|began|increase|increased|increasing|uptitrate|downtitrate|uptitrated|change|changed|titrate|titrated)\b\s+(?<drug>[A-Za-z][A-Za-z\- ]+?)(?:\s+|,|:|;|\b)(?:to\s+)?(?<dose>\d+(?:\.\d+)?)\s*(?<unit>mg)?(?:\s*(?:\/day|per\s*day|mg\/day))?(?:\s*(?<freq>q\.?\s*d\.?|b\.?\s*i\.?\s*d\.?|t\.?\s*i\.?\s*d\.?|q\.?\s*h\.?\s*s\.?|qod|qd|bid|tid|qid|qhs|daily|once\s+daily|twice\s+daily))?/gi;
  const MED_STOP_RE = /\b(stop|stopped|dc|discontinue|discontinued)\b\s+(?<drug>[A-Za-z][A-Za-z\- ]+)/gi;
  const MED_CHANGE_RE = /\b(?<drug>[A-Za-z][A-Za-z\- ]+?)\b\s+(?:was\s+)?(?:increase(?:d)?|uptitrate(?:d)?|up[-\s]?titrate(?:d)?|change(?:d)?\s+to|titrate(?:d)?\s+to)\s+(?<dose>\d+(?:\.\d+)?)\s*(?<unit>mg)?(?:\s*(?:\/day|per\s*day|mg\/day))?(?:\s*(?<freq>q\.?\s*d\.?|b\.?\s*i\.?\s*d\.?|t\.?\s*i\.?\s*d\.?|q\.?\s*h\.?\s*s\.?|qod|qd|bid|tid|qid|qhs|daily|once\s+daily|twice\s+daily))?/gi;
  const MED_BARE_RE = /^(?<drug>[A-Za-z][A-Za-z\-]+)\s+(?<dose>\d+(?:\.\d+)?)\s*(?<unit>mg)(?:\s*(?:\/day|per\s*day|mg\/day))?(?:\s*(?<freq>q\.?\s*d\.?|b\.?\s*i\.?\s*d\.?|t\.?\s*i\.?\s*d\.?|q\.?\s*h\.?\s*s\.?|qod|qd|bid|tid|qid|qhs|daily|once\s+daily|twice\s+daily))?$/i;

  function canonDrug(name) {
    name = (name||'').replace(/\s+/g,' ').trim();
    const n = name.toLowerCase();
    if (['earlier','prior','previous','to','the','a','an','all','my','meds','comment','both','of','these'].includes(n)) return null;
    if (n.length < 3) return null;
    const typo = { valsarstan:'valsartan', valsarsan:'valsartan' };
    const fixed = typo[n]||n;
    return fixed.charAt(0).toUpperCase() + fixed.slice(1);
  }

  function canonFreq(freq) {
    if (!freq) return null;
    let f = freq.trim().toLowerCase().replace(/\./g,'').replace(/\s+/g,' ');
    f = f.replace('q d','qd').replace('b i d','bid').replace('t i d','tid');
    return FREQ_CANON[f]||f;
  }

  function computeDailyDose(doseMg, freq) {
    if (doseMg==null) return null;
    const mult = {qd:1,bid:2,tid:3,qid:4}[freq]||1;
    return doseMg * mult;
  }

  function medSigLabel(med, dose) {
    const nm = med.toLowerCase();
    if (nm==='valsartan') { return dose<=80.001?'80 mg qd':'80 mg bid'; }
    if (nm==='metoprolol') { return `${dose} mg qd`; }
    if (nm==='candesartan') { return `${dose} mg qd`; }
    if (nm==='irbesartan') { return `${dose} mg qd`; }
    return `${dose} mg qd`;
  }

  function parseMedIntervals(readingsAll, startDt, endDt) {
    const events = [];
    for (const r of readingsAll) {
      if (!r.notes) continue;
      const dt = r.t;
      const lines = splitNoteLines(r.notes);

      for (const line of lines) {
        for (const mm of line.matchAll(MED_STOP_RE)) {
          const drug = canonDrug(mm.groups.drug);
          if (!drug) continue;
          events.push({ dt, type:'stop', drug, dose:null, freq:null, daily:null, sig:null });
        }
        for (const mm of line.matchAll(MED_CHANGE_RE)) {
          const drug = canonDrug(mm.groups.drug);
          if (!drug) continue;
          const doseMg = mm.groups.dose ? parseFloat(mm.groups.dose) : null;
          const freq = canonFreq(mm.groups.freq);
          const daily = doseMg!=null ? computeDailyDose(doseMg, freq) : null;
          let sig = null;
          if (doseMg!=null && freq) sig = `${doseMg} mg ${freq}`;
          else if (doseMg!=null) sig = `${doseMg} mg`;
          events.push({ dt, type:'start', drug, dose:doseMg, freq, daily, sig });
        }
        for (const mm of line.matchAll(MED_START_RE)) {
          const drug = canonDrug(mm.groups.drug);
          if (!drug) continue;
          const doseMg = mm.groups.dose ? parseFloat(mm.groups.dose) : null;
          const freq = canonFreq(mm.groups.freq);
          const daily = doseMg!=null ? computeDailyDose(doseMg, freq) : null;
          let sig = null;
          if (doseMg!=null && freq) sig = `${doseMg} mg ${freq}`;
          else if (doseMg!=null) sig = `${doseMg} mg`;
          events.push({ dt, type:'start', drug, dose:doseMg, freq, daily, sig });
        }
        const bareMatch = line.match(MED_BARE_RE);
        if (bareMatch && bareMatch.groups) {
          const drug = canonDrug(bareMatch.groups.drug);
          if (drug) {
            const doseMg = parseFloat(bareMatch.groups.dose);
            const freq = canonFreq(bareMatch.groups.freq);
            const daily = computeDailyDose(doseMg, freq);
            let sig = freq ? `${doseMg} mg ${freq}` : `${doseMg} mg`;
            events.push({ dt, type:'start', drug, dose:doseMg, freq, daily, sig });
          }
        }
      }
    }
    events.sort((a,b) => a.dt - b.dt);

    const seen = new Set();
    const deduped = [];
    for (const e of events) {
      const key = `${e.dt.getTime()}|${e.drug}|${e.type}|${e.daily}`;
      if (seen.has(key)) continue;
      seen.add(key);
      deduped.push(e);
    }

    const active = {};
    const intervals = {};

    function closeInterval(drug, stopDate) {
      if (active[drug]) {
        const {start, daily, sig} = active[drug];
        delete active[drug];
        if (start < stopDate) {
          if (!intervals[drug]) intervals[drug] = [];
          intervals[drug].push([start, stopDate, daily, sig]);
        }
      }
    }

    for (const e of deduped) {
      if (e.type === 'stop') {
        closeInterval(e.drug, e.dt);
      } else {
        if (active[e.drug] && e.daily != null && active[e.drug].daily === e.daily && active[e.drug].sig === e.sig) {
          continue;
        }
        closeInterval(e.drug, e.dt);
        active[e.drug] = { start: e.dt, daily: e.daily, sig: e.sig };
      }
    }

    // Close remaining active intervals at endDt — mark as ongoing
    for (const drug of Object.keys(active)) {
      const {start,daily,sig} = active[drug];
      if (start < endDt) {
        if (!intervals[drug]) intervals[drug]=[];
        intervals[drug].push([start, endDt, daily, sig, true]); // 5th element = ongoing flag
      }
    }

    // Clip to display range (preserve ongoing flag)
    const clipped = {};
    for (const [drug, segs] of Object.entries(intervals)) {
      const out = [];
      for (const seg of segs) {
        const [s,e,daily,sig,ongoing] = seg;
        const ss = startDt && s<startDt ? startDt : s;
        const ee = endDt && e>endDt ? endDt : e;
        if (ee>ss) out.push([ss,ee,daily,sig,!!ongoing]);
      }
      if (out.length) clipped[drug]=out;
    }
    return clipped;
  }

  // =============================================
  // DOSE SHADING
  // =============================================
  function getStartingDose(med, segs) {
    const nm = med.toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
    for (const [key,dose] of Object.entries(STARTING_DOSES)) {
      if (nm.includes(key)) return dose;
    }
    const doses = (segs||[]).map(s=>s[2]).filter(d=>d!=null&&d>0);
    return doses.length ? Math.min(...doses) : 1;
  }

  function alphaForDailyDose(dailyMg, startMg) {
    if (!startMg||startMg<=0) return ALPHA_MIN;
    const ratio = dailyMg/startMg;
    if (ratio<=0) return ALPHA_MIN;
    const frac = Math.min(1, (ratio - 1) / (DOSE_RATIO_MAX - 1));
    return ALPHA_MIN + (ALPHA_MAX - ALPHA_MIN) * Math.max(0, frac);
  }

  function getMedColor(drug) {
    let hash=0;
    const s=drug.toLowerCase();
    for(let i=0;i<s.length;i++)hash=((hash<<5)-hash+s.charCodeAt(i))|0;
    return MED_PALETTE[Math.abs(hash)%MED_PALETTE.length];
  }

  function hexToRgb01(hex) {
    hex=hex.replace('#','');
    return [parseInt(hex.substring(0,2),16)/255, parseInt(hex.substring(2,4),16)/255, parseInt(hex.substring(4,6),16)/255];
  }

  function chooseSigTextColor(barHex, alpha) {
    const [r,g,b] = hexToRgb01(barHex);
    const a = Math.max(0,Math.min(1,alpha));
    const bg = [a*r+(1-a)*CANVAS_BG[0], a*g+(1-a)*CANVAS_BG[1], a*b+(1-a)*CANVAS_BG[2]];
    function srgbToLin(c){return c<=0.04045?c/12.92:((c+0.055)/1.055)**2.4;}
    const lum = 0.2126*srgbToLin(bg[0])+0.7152*srgbToLin(bg[1])+0.0722*srgbToLin(bg[2]);
    return lum<0.15?'rgba(255,255,255,.9)':'rgba(0,0,0,.85)';
  }

  // =============================================
  // PROCESS DATA (main pipeline)
  // Accepts config: { amStart, amEnd }
  // Returns { readings, medIntervals, holdEvents }
  // =============================================
  function processData(rawRows, mapping, config) {
    const readings = buildReadings(rawRows, mapping, config);
    const holdEvents = findHoldEventsFromReadings(readings);
    let medIntervals = {};
    if (readings.length > 0) {
      const startDt = readings[0].t;
      const endDt = new Date(readings[readings.length-1].t.getTime() + 12*3600000);
      medIntervals = parseMedIntervals(readings, startDt, endDt);
    }
    return { readings, medIntervals, holdEvents };
  }

  // =============================================
  // RANGE FILTER
  // Accepts rangeDays: number or 'all'
  // =============================================
  function getRangeFiltered(readings, rangeDays) {
    if (!readings.length) return [];
    if (rangeDays === 'all' || rangeDays === 0) return readings.slice();
    const latest = readings[readings.length-1].t.getTime();
    const cutoff = latest - rangeDays * 86400000;
    return readings.filter(x => x.t.getTime() >= cutoff);
  }

  // =============================================
  // GOAL
  // Accepts goalConfig: { profile, customSys, customDia }
  // =============================================
  function getGoal(goalConfig) {
    const p = goalConfig.profile || 'intl';
    if (p === 'us') return { sys:130, dia:80 };
    if (p === 'custom') return { sys: goalConfig.customSys || 135, dia: goalConfig.customDia || 85 };
    return { sys:135, dia:85 };
  }

  // =============================================
  // COMPUTE METRICS
  // =============================================
  function computeMetrics(rs, goalConfig) {
    const goal = getGoal(goalConfig);
    const all={sys:rs.map(x=>x.sys),dia:rs.map(x=>x.dia)};
    const am=rs.filter(x=>x.window==='Morning');
    const pm=rs.filter(x=>x.window==='Afternoon/Evening');
    const severe=rs.filter(x=>x.severe);
    const hypo=rs.filter(x=>x.hypo);
    const aboveGoal=rs.filter(x=>x.sys>=goal.sys||x.dia>=goal.dia);
    return {
      count:rs.length, amCount:am.length, pmCount:pm.length,
      avgSys:mean(all.sys), avgDia:mean(all.dia), sdSys:stddev(all.sys),
      avgAmSys:mean(am.map(x=>x.sys)), avgAmDia:mean(am.map(x=>x.dia)),
      avgPmSys:mean(pm.map(x=>x.sys)), avgPmDia:mean(pm.map(x=>x.dia)),
      pctGreen:rs.length?100*rs.filter(x=>x.green).length/rs.length:0,
      pctRed:rs.length?100*rs.filter(x=>x.red).length/rs.length:0,
      pctSevere:rs.length?100*severe.length/rs.length:0,
      pctHypo:rs.length?100*hypo.length/rs.length:0,
      severe, hypo, aboveGoalPct:rs.length?100*aboveGoal.length/rs.length:0,
      goal
    };
  }

  // =============================================
  // PHASE / STATISTICS FUNCTIONS
  // =============================================
  function phaseStatsFromReadings(sub) {
    const n = sub.length;
    if (!n) return { meanSBP: null, meanDBP: null, green: null, red: null, severe: null, n: 0 };
    return {
      meanSBP: mean(sub.map(x => x.sys)),
      meanDBP: mean(sub.map(x => x.dia)),
      green: 100 * sub.filter(x => x.green).length / n,
      red: 100 * sub.filter(x => x.red).length / n,
      severe: 100 * sub.filter(x => x.severe).length / n,
      n
    };
  }

  function getLastNReadings(rs, beforeDate, N, lookbackMonths, windowLabel, minPhaseDays) {
    const lookbackMs = lookbackMonths * 30.44 * 86400000;
    const earliest = new Date(beforeDate.getTime() - lookbackMs);
    const candidates = rs.filter(r => r.t < beforeDate && r.t >= earliest && (windowLabel === 'All' || r.window === windowLabel));
    candidates.sort((a, b) => b.t - a.t);
    const byCount = candidates.slice(0, N);
    const minDaysCutoff = new Date(beforeDate.getTime() - minPhaseDays * 86400000);
    const byDays = candidates.filter(r => r.t >= minDaysCutoff);
    const countSpan = byCount.length >= 2 ? (byCount[0].t.getTime() - byCount[byCount.length - 1].t.getTime()) : 0;
    const daySpan = minPhaseDays * 86400000;
    const taken = countSpan >= daySpan ? byCount : byDays;
    taken.sort((a, b) => a.t - b.t);
    return taken;
  }

  function getRecentReadings(rs, weeksBack, windowLabel) {
    if (!rs.length) return [];
    const latest = rs[rs.length - 1].t;
    const cutoff = new Date(latest.getTime() - weeksBack * 7 * 86400000);
    return rs.filter(r => r.t >= cutoff && (windowLabel === 'All' || r.window === windowLabel));
  }

  function getRecentReadingsBeforeDate(rs, beforeDate, weeksBack, windowLabel) {
    if (!rs.length) return [];
    const cutoff = new Date(beforeDate.getTime() - weeksBack * 7 * 86400000);
    return rs.filter(r => r.t >= cutoff && r.t < beforeDate && (windowLabel === 'All' || r.window === windowLabel));
  }

  function phaseStats(rs, startDt, endDt, windowLabel) {
    const sub = rs.filter(r => r.t >= startDt && r.t < endDt && (windowLabel === 'All' || r.window === windowLabel));
    return phaseStatsFromReadings(sub);
  }

  function buildMedChangeBoundaries(readings, medIntervals) {
    if (!readings.length) return [];
    const dataEndTs = readings[readings.length - 1].t.getTime();
    const dateSet = new Set();
    for (const [drug, segs] of Object.entries(medIntervals)) {
      for (const seg of segs) {
        if (seg[0]) dateSet.add(seg[0].getTime());
        if (seg[1] && seg[1].getTime() < dataEndTs) {
          dateSet.add(seg[1].getTime() + 86400000);
        }
      }
    }
    const sortedDates = [...dateSet].sort((a, b) => a - b);
    const boundaries = [];
    let prevLabel = '';
    for (const ts of sortedDates) {
      const activeMeds = [];
      for (const [drug, segs] of Object.entries(medIntervals)) {
        for (const seg of segs) {
          const [s, e, daily] = seg;
          if (!s || !e) continue;
          if (s.getTime() <= ts && e.getTime() >= ts) {
            const sigLabel = daily != null ? `${drug} ${medSigLabel(drug, daily)}` : drug;
            activeMeds.push(sigLabel);
            break;
          }
        }
      }
      const label = activeMeds.join(' + ');
      if (label && label !== prevLabel) {
        boundaries.push({ date: new Date(ts), changeLabel: label });
        prevLabel = label;
      }
    }
    const merged = [];
    for (const b of boundaries) {
      if (merged.length && (b.date.getTime() - merged[merged.length - 1].date.getTime()) < 2 * 86400000) {
        merged[merged.length - 1] = { ...b };
      } else {
        merged.push({ ...b });
      }
    }
    return merged;
  }

  function getCurrentMeds(readings, medIntervals) {
    const current = [];
    for (const [drug, segs] of Object.entries(medIntervals)) {
      if (!segs.length) continue;
      const lastSeg = segs[segs.length - 1];
      const [start, end, daily, sig] = lastSeg;
      if (readings.length > 0) {
        const dataEnd = readings[readings.length - 1].t;
        if (end.getTime() >= dataEnd.getTime() - 86400000) {
          const label = daily != null ? `${drug} ${medSigLabel(drug, daily)}` : drug;
          current.push(label);
        }
      }
    }
    return current;
  }

  // =============================================
  // COMPARISON HELPERS
  // =============================================
  function bpTier(dSBP, dDBP) {
    if (dSBP == null || dDBP == null) return 'none';
    const maxDrop = Math.max(dSBP, dDBP);
    const minDrop = Math.min(dSBP, dDBP);
    if (maxDrop <= -10) return 'significant';
    if (maxDrop <= -5 || minDrop <= -5) return 'moderate';
    if (maxDrop < 0 || minDrop < 0) return 'mild';
    if (maxDrop <= 3 && minDrop >= -3) return 'flat';
    return 'increase';
  }

  function bpResponseLabel(dSBP, dDBP) {
    const tier = bpTier(dSBP, dDBP);
    const labels = {
      'significant': '▼▼ Significant improvement',
      'moderate':    '▼ Moderate improvement',
      'mild':        '~ Mild improvement',
      'flat':        '— Flat / No meaningful change',
      'increase':    '▲ BP increased',
      'none':        '—'
    };
    return labels[tier] || '—';
  }

  function bpArrow(dSBP, dDBP) {
    const tier = bpTier(dSBP, dDBP);
    const arrows = { 'significant':'⬇️','moderate':'↓','mild':'↘','flat':'→','increase':'↑','none':'' };
    return arrows[tier] || '';
  }

  function ppLabel(dPP, type) {
    if (dPP == null) return '';
    const abs = Math.abs(dPP);
    const dir = dPP < 0 ? 'narrowed' : 'widened';
    const unit = type === 'pp' ? 'mmHg' : 'bpm';
    if (abs < 2) return `${type === 'pp' ? 'PP' : 'HR'} stable`;
    return `${type === 'pp' ? 'PP' : 'HR'} ${dir} ${abs.toFixed(0)} ${unit}`;
  }

  // =============================================
  // LOCAL STORAGE PERSISTENCE
  // =============================================
  function saveToLocalStorage(name, base64, prefix) {
    const p = prefix || 'hbp';
    try { localStorage.setItem(p + '_fileName', name); localStorage.setItem(p + '_fileData', base64); } catch(e) { console.warn('localStorage save failed:', e); }
  }
  function loadFromLocalStorage(prefix) {
    const p = prefix || 'hbp';
    try {
      const name = localStorage.getItem(p + '_fileName');
      const data = localStorage.getItem(p + '_fileData');
      if (name && data) return { name, data };
    } catch(e) {}
    return null;
  }
  function arrayBufferToBase64(buf) {
    const bytes = new Uint8Array(buf);
    let bin = '';
    for (let i=0; i<bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    return btoa(bin);
  }
  function base64ToArrayBuffer(b64) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i=0; i<bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes.buffer;
  }

  // =============================================
  // PUBLIC API — everything both versions need
  // =============================================
  return {
    // Constants
    COLOR_GREEN, COLOR_RED, COLOR_SEVERE, COLOR_HYPO, HR_COLOR, TARGET_COLOR, HOLD_STAR_COLOR,
    TARGET_SBP, TARGET_DBP, GREEN_SBP_MAX, GREEN_DBP_MAX,
    SEVERE_SBP_MIN, SEVERE_DBP_MIN, SEVERE_LABEL_SBP_MIN,
    HYPO_SBP_MAX, HYPO_DBP_MAX,
    BASE_BAR_W, SEVERE_BAR_W, MED_PALETTE, CANVAS_BG,
    ALPHA_MIN, ALPHA_MAX, DOSE_RATIO_MAX, STARTING_DOSES, FREQ_CANON,

    // Utilities
    mean, stddev, quantile, fmtBP, fmtDate, fmtDateShort, fmtPct, fmtNum, daysBetween,

    // Parsing
    parseCSV, normalizeHeader, sheetToAoa, findLikelyHeaderRow, excelAoaToRows,
    tryParseDate, parseDateAndTime, detectColumns,

    // Classification
    classifyBP, reclassifyForGoal,

    // Data pipeline
    buildReadings, findHoldEventsFromReadings, processData,

    // Medication engine
    splitNoteLines, canonDrug, canonFreq, computeDailyDose, medSigLabel,
    parseMedIntervals,

    // Dose shading
    getStartingDose, alphaForDailyDose, getMedColor, hexToRgb01, chooseSigTextColor,

    // Filtering & metrics
    getRangeFiltered, getGoal, computeMetrics,

    // Phase / statistics
    phaseStatsFromReadings, getLastNReadings, getRecentReadings, getRecentReadingsBeforeDate, phaseStats,
    buildMedChangeBoundaries, getCurrentMeds,

    // Comparison helpers
    bpTier, bpResponseLabel, bpArrow, ppLabel,

    // Persistence
    saveToLocalStorage, loadFromLocalStorage, arrayBufferToBase64, base64ToArrayBuffer,
  };
})();
