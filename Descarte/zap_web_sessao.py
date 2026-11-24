#!/usr/bin/env python3
"""Dump/restore portável da sessão WhatsApp Web (Windows -> Linux)."""

import json
import os

from cryptography.fernet import Fernet

FORMATO_PORTAVEL = "zap-portable-v1"

# Serializa ArrayBuffer/Date para JSON; executado no Chromium via page.evaluate.
_JS_CODEC = r"""
function zapEnc(v) {
  if (v === null || v === undefined) return v;
  const t = typeof v;
  if (t === 'string' || t === 'number' || t === 'boolean') return v;
  if (v instanceof Date) return {__zap: 'Date', v: v.toISOString()};
  if (typeof ArrayBuffer !== 'undefined' && v instanceof ArrayBuffer) {
    return {__zap: 'AB', v: zapB64(new Uint8Array(v))};
  }
  if (typeof ArrayBuffer !== 'undefined' && ArrayBuffer.isView(v)) {
    return {__zap: 'U8', n: v.constructor.name, v: zapB64(new Uint8Array(v.buffer, v.byteOffset, v.byteLength))};
  }
  if (Array.isArray(v)) return v.map(zapEnc);
  if (t === 'object') {
    const o = {};
    for (const k of Object.keys(v)) o[k] = zapEnc(v[k]);
    return o;
  }
  return null;
}
function zapB64(u8) {
  let s = '';
  const chunk = 0x8000;
  for (let i = 0; i < u8.length; i += chunk) {
    s += String.fromCharCode.apply(null, u8.subarray(i, i + chunk));
  }
  return btoa(s);
}
function zapU8(b64) {
  const bin = atob(b64);
  const u8 = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
  return u8;
}
function zapDec(v) {
  if (v === null || v === undefined) return v;
  if (Array.isArray(v)) return v.map(zapDec);
  if (v && typeof v === 'object' && v.__zap === 'Date') return new Date(v.v);
  if (v && typeof v === 'object' && v.__zap === 'AB') return zapU8(v.v).buffer;
  if (v && typeof v === 'object' && v.__zap === 'U8') {
    const u8 = zapU8(v.v);
    if (v.n === 'Uint8Array') return u8;
    return u8;
  }
  if (v && typeof v === 'object') {
    const o = {};
    for (const k of Object.keys(v)) o[k] = zapDec(v[k]);
    return o;
  }
  return v;
}
"""

DUMP_JS = """
async () => {
  %s
  const ls = {};
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      ls[k] = localStorage.getItem(k);
    }
  } catch (e) {}

  const idb = [];
  let dbs = [];
  try {
    dbs = await indexedDB.databases();
  } catch (e) { dbs = []; }

  function dumpDb(name, version) {
    return new Promise((resolve) => {
      let timedOut = false;
      const timer = setTimeout(() => { timedOut = true; resolve(null); }, 20000);
      let req;
      try {
        req = version ? indexedDB.open(name, version) : indexedDB.open(name);
      } catch (e) {
        clearTimeout(timer);
        resolve(null);
        return;
      }
      req.onerror = () => { clearTimeout(timer); resolve(null); };
      req.onupgradeneeded = (ev) => {
        try { ev.target.transaction.abort(); } catch (e) {}
      };
      req.onsuccess = () => {
        if (timedOut) { try { req.result.close(); } catch (e) {} return; }
        const db = req.result;
        const storeNames = Array.from(db.objectStoreNames);
        const out = {name: db.name, version: db.version, stores: {}};
        if (!storeNames.length) {
          db.close();
          clearTimeout(timer);
          resolve(out);
          return;
        }
        const tx = db.transaction(storeNames, 'readonly');
        let pending = storeNames.length;
        for (const storeName of storeNames) {
          const store = tx.objectStore(storeName);
          const meta = {
            keyPath: store.keyPath,
            autoIncrement: !!store.autoIncrement,
            indexes: Array.from(store.indexNames).map((n) => {
              const idx = store.index(n);
              return {name: idx.name, keyPath: idx.keyPath, unique: !!idx.unique, multiEntry: !!idx.multiEntry};
            }),
            records: []
          };
          const kr = store.openCursor();
          kr.onerror = () => {
            out.stores[storeName] = meta;
            pending -= 1;
            if (pending === 0) { db.close(); clearTimeout(timer); resolve(out); }
          };
          kr.onsuccess = (ev) => {
            const cursor = ev.target.result;
            if (cursor && meta.records.length < 8000) {
              meta.records.push({key: zapEnc(cursor.key), value: zapEnc(cursor.value)});
              cursor.continue();
            } else {
              out.stores[storeName] = meta;
              pending -= 1;
              if (pending === 0) { db.close(); clearTimeout(timer); resolve(out); }
            }
          };
        }
        tx.onerror = () => { try { db.close(); } catch (e) {} clearTimeout(timer); resolve(out); };
      };
    });
  }

  for (const info of dbs) {
    if (!info || !info.name) continue;
    const dumped = await dumpDb(info.name, info.version);
    if (dumped) idb.push(dumped);
  }

  return {
    localStorage: ls,
    indexedDB: idb,
    userAgent: navigator.userAgent,
    platform: navigator.platform
  };
}
""" % _JS_CODEC

RESTORE_IDB_JS = """
async (idbList) => {
  %s
  async function deleteDb(name) {
    return new Promise((resolve) => {
      try {
        const req = indexedDB.deleteDatabase(name);
        const done = () => resolve();
        req.onsuccess = done;
        req.onerror = done;
        req.onblocked = done;
        setTimeout(done, 5000);
      } catch (e) { resolve(); }
    });
  }
  async function importDb(data) {
    await deleteDb(data.name);
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(data.name, data.version || 1);
      req.onerror = () => reject(req.error || new Error('open failed'));
      req.onupgradeneeded = () => {
        const db = req.result;
        for (const [storeName, store] of Object.entries(data.stores || {})) {
          if (db.objectStoreNames.contains(storeName)) db.deleteObjectStore(storeName);
          const opts = {};
          if (store.keyPath !== undefined && store.keyPath !== null) opts.keyPath = store.keyPath;
          if (store.autoIncrement) opts.autoIncrement = true;
          const os = db.createObjectStore(storeName, opts);
          for (const idx of (store.indexes || [])) {
            try {
              os.createIndex(idx.name, idx.keyPath, {unique: !!idx.unique, multiEntry: !!idx.multiEntry});
            } catch (e) {}
          }
        }
      };
      req.onsuccess = () => {
        const db = req.result;
        const names = Object.keys(data.stores || {});
        if (!names.length) { db.close(); resolve(); return; }
        const existing = names.filter((n) => db.objectStoreNames.contains(n));
        if (!existing.length) { db.close(); resolve(); return; }
        const tx = db.transaction(existing, 'readwrite');
        tx.oncomplete = () => { db.close(); resolve(); };
        tx.onerror = () => { try { db.close(); } catch (e) {} reject(tx.error); };
        for (const storeName of existing) {
          const store = data.stores[storeName];
          const os = tx.objectStore(storeName);
          for (const rec of (store.records || [])) {
            try {
              const value = zapDec(rec.value);
              const key = rec.key !== undefined ? zapDec(rec.key) : undefined;
              if (store.keyPath === null || store.keyPath === undefined) {
                os.put(value, key);
              } else {
                os.put(value);
              }
            } catch (e) {}
          }
        }
      };
    });
  }
  const ok = [];
  const fail = [];
  for (const db of (idbList || [])) {
    try {
      await importDb(db);
      ok.push(db.name);
    } catch (e) {
      fail.push(db.name + ': ' + (e && e.message ? e.message : String(e)));
    }
  }
  return {ok, fail};
}
""" % _JS_CODEC


def eh_portavel(payload):
    return isinstance(payload, dict) and payload.get("format") == FORMATO_PORTAVEL


def descriptografar_sessao(caminho_enc, chave):
    f = Fernet(chave.encode() if isinstance(chave, str) else chave)
    with open(caminho_enc, "rb") as fp:
        dados = f.decrypt(fp.read())
    if dados[:1] == b"{":
        return json.loads(dados.decode("utf-8"))
    return {"format": "chromium-zip", "zip_bytes": dados}


def criptografar_payload(payload, chave):
    f = Fernet(chave.encode() if isinstance(chave, str) else chave)
    bruto = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return f.encrypt(bruto)


def obter_ou_criar_chave(caminho_key):
    if os.path.exists(caminho_key):
        with open(caminho_key, "r", encoding="utf-8") as fp:
            existente = fp.read().strip()
        if existente:
            return existente.encode(), False
    chave = Fernet.generate_key()
    with open(caminho_key, "w", encoding="utf-8") as fp:
        fp.write(chave.decode())
    return chave, True


def script_stealth_e_localstorage(local_storage, platform="Win32"):
    ls_json = json.dumps(local_storage or {}, ensure_ascii=False)
    plat_json = json.dumps(platform or "Win32")
    return f"""
(() => {{
  try {{
    Object.defineProperty(navigator, 'webdriver', {{get: () => undefined}});
    Object.defineProperty(navigator, 'platform', {{get: () => {plat_json}}});
    Object.defineProperty(navigator, 'languages', {{get: () => ['pt-BR', 'pt', 'en-US', 'en']}});
  }} catch (e) {{}}
  try {{
    const ls = {ls_json};
    for (const k of Object.keys(ls)) {{
      try {{ localStorage.setItem(k, ls[k]); }} catch (e) {{}}
    }}
  }} catch (e) {{}}
}})();
"""


def dump_sessao_pagina(page, context):
    storage_state = context.storage_state()
    extra = page.evaluate(DUMP_JS)
    payload = {
        "format": FORMATO_PORTAVEL,
        "storage_state": storage_state,
        "localStorage": extra.get("localStorage") or {},
        "indexedDB": extra.get("indexedDB") or [],
        "userAgent": extra.get("userAgent") or "",
        "platform": extra.get("platform") or "Win32",
    }
    # Garante localStorage também em storage_state.origins (Playwright)
    origem = "https://web.whatsapp.com"
    origins = payload["storage_state"].setdefault("origins", [])
    ls_items = [{"name": k, "value": v} for k, v in payload["localStorage"].items() if v is not None]
    found = False
    for o in origins:
        if o.get("origin") == origem:
            o["localStorage"] = ls_items
            found = True
            break
    if not found and ls_items:
        origins.append({"origin": origem, "localStorage": ls_items})
    return payload


def restaurar_indexeddb(page, idb_list):
    if not idb_list:
        return {"ok": [], "fail": ["indexedDB vazio"]}
    return page.evaluate(RESTORE_IDB_JS, idb_list)
