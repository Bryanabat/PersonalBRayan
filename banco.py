# banrep_flujo_completo.py
import argparse
import re
import time
import unicodedata

from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

URL_DEFAULT = (
    "https://suameca.banrep.gov.co/estadisticas-economicas-back/"
    "reporte-oac.html?path=%2FEstadisticas_Banco_de_la_Republica%2F4_Sector_Externo_tasas_de_cambio_y_derivados"
    "%2F1_Tasas_de_cambio%2F2_Tasas_de_cambio_otros_paises%2F3_Tasas_de_cambio_paises_del_mundo_Todas_las_monedas_old"
)

TARGETS = [
    "Corona sueca",
    "Dólar canadiense",
    "Dólar australiano",
    "Euro",
    "Franco suizo",
    "Peso chileno",
    "Libra esterlina",
    "Real brasileño",
    "Yen japonés",
]

# Códigos ISO esperados (para cruzar por código si el nombre no coincide exactamente)
CODE_BY_NAME = {
    "corona sueca": "SEK",
    "dolar canadiense": "CAD",
    "dolar australiano": "AUD",
    "euro": "EUR",
    "franco suizo": "CHF",
    "peso chileno": "CLP",
    "libra esterlina": "GBP",
    "real brasileno": "BRL",
    "yen japones": "JPY",
}

# ================
# Utilidades base
# ================
def build_driver():
    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-features=msEdgeBackgroundTabSuspension,msEdgeLazyLoad,WebUsb")
    opts.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    opts.add_experimental_option("detach", True)  # NO cerrar Edge al finalizar
    opts.add_argument("--log-level=3")
    return webdriver.Edge(options=opts)

def wait_until_ready(driver, timeout=120):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("""
            const dvReady=[...document.querySelectorAll('oracle-dv')]
              .some(el=>el.classList && el.classList.contains('oj-complete'));
            const anyFilter=document.querySelector('[id^="dashboardfilterviz_box_"] .bi_dashboardfilterviz_tile_wrapper');
            return dvReady || !!anyFilter;
        """)
    )

def aceptar_cookies(driver):
    try:
        btn = WebDriverWait(driver, 6).until(
            EC.element_to_be_clickable((By.XPATH,
                "//*[self::button or self::a or self::span]"
                "[contains(translate(., 'ACEPTCONSENTDEACUERDOOK', 'aceptconsentdeacuerdook'),'acept')]"
            ))
        )
        try: btn.click()
        except Exception: driver.execute_script("arguments[0].click();", btn)
    except Exception:
        pass

def open_filter_tile(driver, tile_id, timeout=40):
    tile = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, f"#{tile_id} .bi_dashboardfilterviz_tile_wrapper"))
    )
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", tile)
    try: tile.click()
    except Exception: driver.execute_script("arguments[0].click();", tile)

def click_shuttle_option_and_add(driver, option_text, timeout=40, tries=3):
    wait = WebDriverWait(driver, timeout)

    def _find_option():
        els = driver.find_elements(
            By.CSS_SELECTOR,
            f".biShuttleAvailableValuesItem[data-bi-shuttle-display-value='{option_text}']"
        )
        if els: return els[0]
        return wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//*[contains(@class,'biShuttleAvailableValuesItem') and "
            f"(normalize-space()='{option_text}' or contains(normalize-space(), '{option_text}'))]"
        )))

    def _js_click(el):
        driver.execute_script("""
            const el = arguments[0];
            try { el.scrollIntoView({block:'center'}); } catch(e){}
            for (const t of ['mouseover','mousemove','mousedown','mouseup','click']) {
              el.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,view:window}));
            }
        """, el)

    def _is_in_selections():
        return driver.execute_script("""
            const txt = arguments[0];
            const right = [...document.querySelectorAll('*')]
              .find(n => /Selecciones/i.test(n.textContent||'') &&
                         n.closest && n.closest('[class*="oj-"]'));
            if (!right) return false;
            return (right.textContent||'').trim().includes(txt);
        """, option_text) is True

    last_err = None
    for _ in range(tries):
        try:
            opt = _find_option()
            _js_click(opt)
            driver.implicitly_wait(0)
            for _ in range(10):
                if _is_in_selections(): return
                driver.execute_script("return 0")
            driver.implicitly_wait(5)
            try:
                add_btn = wait.until(EC.element_to_be_clickable((
                    By.XPATH, "//span[contains(normalize-space(),'Agregar')]/ancestor::button[1]"
                )))
                try: add_btn.click()
                except Exception: driver.execute_script("arguments[0].click();", add_btn)
            except TimeoutException:
                pass
            if _is_in_selections(): return
        except StaleElementReferenceException as e:
            last_err = e
            continue
        except TimeoutException as e:
            raise e
    if last_err: raise last_err
    raise TimeoutException(f"No se pudo seleccionar '{option_text}' en el shuttle.")

def click_shuttle_option_only(driver, option_text, timeout=40, tries=3):
    wait = WebDriverWait(driver, timeout)
    def _find_option():
        els = driver.find_elements(
            By.CSS_SELECTOR,
            f".biShuttleAvailableValuesItem[data-bi-shuttle-display-value='{option_text}']"
        )
        if els: return els[0]
        return wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//*[contains(@class,'biShuttleAvailableValuesItem') and normalize-space()="
            f"'{option_text}']"
        )))
    last_err = None
    for _ in range(tries):
        try:
            opt = _find_option()
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", opt)
            try: opt.click()
            except Exception: driver.execute_script("arguments[0].click();", opt)
            WebDriverWait(driver, 2).until(lambda d: True)
            return
        except StaleElementReferenceException as e:
            last_err = e
            continue
        except TimeoutException as e:
            raise e
    if last_err: raise last_err

# =========================
# Fecha (JS helpers)
# =========================
JS_FN_SELECT_FROM_FILTER = r"""
function __selectFromFilter(label, optionText){
  const getVisibleContainer = () => {
    const cand = [...document.querySelectorAll('.oj-popup, .oj-dialog, [role="dialog"]')]
      .filter(el => getComputedStyle(el).display!=='none' && el.offsetParent!==null);
    return cand[0] || document;
  };

  function openCombo(label){
    const root = getVisibleContainer();
    const combo = root.querySelector('div[role="combobox"][aria-label="'+label+'"]')
               || document.querySelector('div[role="combobox"][aria-label="'+label+'"]');
    if (!combo) return {ok:false, step:'combo-not-found', label};
    let baseId = null;
    const chosen = combo.querySelector('span.oj-select-chosen[id$="_selected"]');
    if (chosen && chosen.id && chosen.id.endsWith('_selected')) baseId = chosen.id.slice(0,-9);
    else baseId = (combo.id || '').replace(/^oj-select-choice-/, '') || null;

    try {
      if (window.jQuery && jQuery.fn && typeof jQuery.fn.ojSelect==='function' && baseId) {
        jQuery('#'+baseId).ojSelect('open'); return {ok:true, step:'opened-jet', combo};
      }
    } catch(e){}
    try {
      if (baseId) {
        const host = document.getElementById(baseId);
        if (host) {
          let el = host.closest('oj-select-one,oj-select-single,oj-combobox-one');
          if (el && typeof el.open==='function') { el.open(); return {ok:true, step:'opened-ce', combo}; }
        }
      }
    } catch(e){}
    try {
      if (chosen) { ['mousedown','mouseup','click'].forEach(t=>chosen.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true}))); return {ok:true, step:'opened-text', combo}; }
    } catch(e){}
    try {
      const arrow = combo.querySelector('.oj-select-open-icon, .oj-select-arrow');
      if (arrow) { ['mousedown','mouseup','click'].forEach(t=>arrow.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true}))); return {ok:true, step:'opened-arrow', combo}; }
    } catch(e){}
    try { combo.focus(); combo.dispatchEvent(new KeyboardEvent('keydown',{key:'ArrowDown',altKey:true,bubbles:true})); return {ok:true, step:'opened-alt-down', combo}; } catch(e){}
    return {ok:false, step:'open-failed', label};
  }

  function getListbox(){
    const uls = [...document.querySelectorAll('div[id^="oj-listbox-drop"] ul.oj-listbox-results')]
                .filter(ul => getComputedStyle(ul).display!=='none');
    return uls[0] || null;
  }

  function clickByExactText(root, txt){
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT);
    let n; const norm = s => (s||'').trim();
    while ((n = walker.nextNode())) {
      if (norm(n.textContent) === txt) {
        const hit = n.querySelector('div,span,a,label') || n;
        ['mousedown','mouseup','click'].forEach(t=>hit.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true})));
        return true;
      }
    }
    return false;
  }

  const opened = openCombo(label);
  if (!opened.ok) return opened;

  return new Promise(async (resolve)=>{
    let lb = null;
    for (let i=0;i<25;i++){ lb = getListbox(); if (lb) break; await new Promise(r=>setTimeout(r,200)); }
    if (!lb) return resolve({ok:false, step:'listbox-timeout', label});

    if (!clickByExactText(lb, optionText)) {
      return resolve({ok:false, step:'option-not-found', label, option: optionText});
    }

    const cont = getVisibleContainer();
    const addBtn = [...cont.querySelectorAll('button, a')].find(b => /agregar/i.test(b.textContent || '') || /agregar/i.test(b.getAttribute('aria-label') || ''));
    if (addBtn) {
      ['mousedown','mouseup','click'].forEach(t=>addBtn.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true})));
    }

    resolve({ok:true, step:'done', label, option: optionText});
  });
}
"""

JS_FN_SET_FECHA = r"""
function __setFecha(comparatorText, dateStr){
  const root = (()=>{
    const cand = [...document.querySelectorAll('.oj-popup, .oj-dialog, [role="dialog"]')]
      .filter(el => getComputedStyle(el).display!=='none' && el.offsetParent!==null);
    return cand[0] || document;
  })();

  if (typeof __selectFromFilter !== 'function'){
    %s
  }

  return __selectFromFilter("Tipo de Rango", comparatorText).then(function(res){
    if (!res || !res.ok) return res || {ok:false, step:'fecha-comparator-failed'};

    const input = root.querySelector('input.oj-inputdatetime-input') ||
                  root.querySelector('input.oj-inputtext-input');
    if (!input) return {ok:false, step:'fecha-input-not-found'};

    input.focus();
    input.value = "";
    input.dispatchEvent(new Event('input',{bubbles:true}));
    input.value = dateStr;
    input.dispatchEvent(new Event('input',{bubbles:true}));
    input.dispatchEvent(new Event('change',{bubbles:true}));
    input.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));
    input.dispatchEvent(new KeyboardEvent('keyup',{key:'Enter',bubbles:true}));
    input.blur();

    return {ok:true, step:'fecha-done', comparator: comparatorText, date: dateStr};
  });
}
""" % JS_FN_SELECT_FROM_FILTER

def call_js_function(driver, fn_code, fn_name, *fn_args):
    wrapper = f"""
      var cb = arguments[arguments.length-1];
      var a = Array.prototype.slice.call(arguments,0,arguments.length-1);
      {fn_code}
      var __res;
      try {{
        __res = {fn_name}.apply(null, a);
      }} catch(e) {{
        cb({{ok:false, step:'js-throw', error:String(e)}}); return;
      }}
      if (__res && typeof __res.then === 'function') {{
        __res.then(function(r){{ cb(r); }})
             .catch(function(e){{ cb({{ok:false, step:'js-error', error:String(e)}}); }});
      }} else {{
        cb(__res);
      }}
    """
    return driver.execute_async_script(wrapper, *fn_args)

# =========================================================
# Iframe + esperar grid
# =========================================================
def switch_to_frame_with_selector(driver, css_selector, max_depth=5):
    driver.switch_to.default_content()

    def doc_has_selector(drv):
        try:
            return bool(drv.execute_script(
                "return !!document.querySelector(arguments[0]);", css_selector
            ))
        except Exception:
            return False

    if doc_has_selector(driver):
        return True

    def dfs(level=1):
        if level > max_depth:
            return False
        frames = driver.find_elements(By.CSS_SELECTOR, "iframe")
        for f in frames:
            if not f.is_displayed():
                continue
            try:
                driver.switch_to.frame(f)
                if doc_has_selector(driver):
                    return True
                if dfs(level+1):
                    return True
            except Exception:
                pass
            finally:
                driver.switch_to.parent_frame()
        return False

    return dfs()

def wait_for_grid_loaded(driver, timeout=60):
    ok = switch_to_frame_with_selector(driver, "oj-data-grid", max_depth=6)
    if not ok:
        raise RuntimeError("No encontré el <oj-data-grid> en ningún iframe.")

    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return !!document.querySelector('oj-data-grid');") is True
    )

    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("""
            const grid=document.querySelector('oj-data-grid');
            if(!grid) return false;
            const h0=grid.querySelectorAll('div.oj-datagrid-header-grouping[data-oj-level="0"] > div.oj-datagrid-header-cell').length
                   || document.querySelectorAll('div.oj-datagrid-header-grouping[data-oj-level="0"] > div.oj-datagrid-header-cell').length;
            const cells=(document.querySelectorAll('div.oj-datagrid-databody div.oj-datagrid-cell span').length
                      || grid.querySelectorAll('div.oj-datagrid-databody div.oj-datagrid-cell span').length);
            return h0>0 && cells>0;
        """) is True
    )

# ============
# Barrido (scroll) del grid y captura de TODAS las columnas
# ============
def sweep_and_read_all_columns(driver, settle_ms=120):
    """
    Recorre horizontalmente el databody, capturando:
    - headers nivel 0 (nombre) y nivel 1 (código), ordenados por 'left'
    - celdas de la primera fila, ordenadas por 'left'
    Devuelve lista de dicts: {'name','code','value'}
    """
    # Llevar al iframe que contiene el grid (si hay)
    switch_to_frame_with_selector(driver, "oj-data-grid", max_depth=6)

    # Referencias
    body = driver.execute_script("return document.querySelector('div[id$=\"OJDataGrid:databody\"]');")
    if not body:
        raise RuntimeError("No encontré el databody del oj-data-grid.")

    # Funciones JS auxiliares
    js_snapshot = r"""
      const grid = document.querySelector('oj-data-grid');
      const h0 = Array.from(document.querySelectorAll('div.oj-datagrid-header-grouping[data-oj-level="0"] > div.oj-datagrid-header-cell'));
      const h1 = Array.from(document.querySelectorAll('div.oj-datagrid-header-grouping[data-oj-level="1"] > div.oj-datagrid-header-cell'));
      const cells = Array.from(document.querySelectorAll('div.oj-datagrid-databody div.oj-datagrid-cell'));

      function pick(items){
        return items.map(div=>{
          const s=div.getAttribute('style')||'';
          const m=/left:\s*([0-9.]+)px/i.exec(s);
          const left=m?parseFloat(m[1]):0;
          return {left,leftKey:String(Math.round(left)), text:(div.textContent||'').trim()};
        });
      }
      return {
        h0: pick(h0),
        h1: pick(h1),
        cells: pick(cells)
      };
    """

    # Ajustes de scroll
    driver.execute_script("const db=document.querySelector('div[id$=\"OJDataGrid:databody\"]'); if(db) db.scrollLeft=0;")
    time.sleep(settle_ms/1000.0)

    # Acumuladores por clave "left" redondeada
    h0_map, h1_map, cell_map = {}, {}, {}

    def merge_snapshot(snap):
        for it in snap["h0"]:
            h0_map.setdefault(it["leftKey"], it)
        for it in snap["h1"]:
            h1_map.setdefault(it["leftKey"], it)
        for it in snap["cells"]:
            cell_map.setdefault(it["leftKey"], it)

    # Primera foto
    snap = driver.execute_script(js_snapshot)
    merge_snapshot(snap)

    # Scroll horizontal hasta el final
    scroll_w = driver.execute_script("const db=document.querySelector('div[id$=\"OJDataGrid:databody\"]'); return db?db.scrollWidth:0;")
    client_w = driver.execute_script("const db=document.querySelector('div[id$=\"OJDataGrid:databody\"]'); return db?db.clientWidth:0;")
    if not scroll_w or not client_w:
        raise RuntimeError("No pude medir scrollWidth/clientWidth del grid.")

    max_left = max(0, scroll_w - client_w)
    step = max(40, int(client_w * 0.85))  # paso de ~85% del viewport

    cur = 0
    while cur < max_left - 1:
        cur = min(cur + step, max_left)
        driver.execute_script("const db=document.querySelector('div[id$=\"OJDataGrid:databody\"]'); if(db) db.scrollLeft=arguments[0];", cur)
        time.sleep(settle_ms/1000.0)  # dar tiempo a render virtualizado
        snap = driver.execute_script(js_snapshot)
        merge_snapshot(snap)

    # Unir por posición
    keys = sorted({*h0_map.keys(), *cell_map.keys()}, key=lambda k: float(k))
    out = []
    for k in keys:
        name = (h0_map.get(k, {}).get("text") or "").strip()
        code = (h1_map.get(k, {}).get("text") or "").strip()
        val  = (cell_map.get(k, {}).get("text") or "").strip()
        if name or val:
            out.append({"name": name, "code": code, "value": val})
    return out

# =========================
# Normalización y matching
# =========================
def _norm(s):
    if s is None: return ""
    s = s.strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

def _similar(a, b):
    A = set(_norm(a).split())
    B = set(_norm(b).split())
    if not A or not B:
        return 0.0
    return len(A & B) / float(len(A | B))

def extraer_objetivo(rows, buscadas=TARGETS):
    idx_by_name = { _norm(r["name"]): r for r in rows }
    idx_by_code = {}
    for r in rows:
        c = (r.get("code") or "").strip().upper()
        if c: idx_by_code[c] = r

    out = []
    for nombre in buscadas:
        n = _norm(nombre)
        r = idx_by_name.get(n)

        if not r:
            exp_code = CODE_BY_NAME.get(n, "")
            if exp_code:
                r = idx_by_code.get(exp_code)

        if not r:
            best = None; best_s = 0.0
            for cand in rows:
                s = _similar(nombre, cand["name"])
                if s > best_s: best_s, best = s, cand
            if best and best_s >= 0.5: r = best

        if r and r.get("value"):
            out.append({"moneda": nombre, "codigo": r.get("code",""), "venta": r["value"]})
        else:
            out.append({"moneda": nombre, "codigo": r.get("code","") if r else "", "venta": "ERROR"})
    return out

# =========================
# Flujo principal
# =========================
def main(url, fecha, comparador, texto_tasa, texto_cambio, espera):
    driver = build_driver()
    driver.get(url)

    wait_until_ready(driver)
    aceptar_cookies(driver)

    # 1) FECHA
    open_filter_tile(driver, "dashboardfilterviz_box_0")
    res_fecha = call_js_function(driver, JS_FN_SET_FECHA, "__setFecha", comparador, fecha)
    print("[Fecha]", res_fecha)

    # 2) TIPO DE TASA
    open_filter_tile(driver, "dashboardfilterviz_box_2")
    click_shuttle_option_and_add(driver, texto_tasa)
    print("[Tipo de Tasa] OK ->", texto_tasa)

    # 3) TIPO DE CAMBIO
    open_filter_tile(driver, "dashboardfilterviz_box_3")
    click_shuttle_option_only(driver, texto_cambio)
    print("[Tipo de Cambio] OK ->", texto_cambio)

    # Espera fija extra para que el grid se reprocese
    if espera and espera > 0:
        print(f"[Espera] {espera:.1f}s para que el grid termine de renderizar…")
        time.sleep(espera)

    # 4) Esperar grid y leer TODO el ancho
    wait_for_grid_loaded(driver, timeout=60)
    rows = sweep_and_read_all_columns(driver, settle_ms=120)

    print("\n[DEBUG] Columnas capturadas:", len(rows))
    for r in rows[:12]:
        print(" -", r["name"], "|", r["code"], "|", r["value"])

    resultados = extraer_objetivo(rows, TARGETS)

    print("\n=== Venta por moneda (tabla visible o virtualizada) ===")
    for r in resultados:
        cod = f" ({r['codigo']})" if r['codigo'] else ""
        print(f"{r['moneda']:<22}{cod:<14}  {r['venta']}")
    print("=======================================================\n")

    input("Listo. Revisa los valores en consola. Presiona ENTER para terminar (Edge queda abierto por 'detach').\n")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", default=URL_DEFAULT)
    parser.add_argument("--fecha", default="20/09/2025", help="dd/mm/yyyy")
    parser.add_argument("--comparador", default="Igual que", choices=["Igual que", "Iniciar en"])
    parser.add_argument("--tasa", default="VENTA", help="Opción exacta del filtro 'Tipo de Tasa'")
    parser.add_argument("--cambio", default="Dólares estadounidenses por cada moneda",
                        help="Opción exacta del filtro 'Tipo de Cambio'")
    parser.add_argument("--espera", type=float, default=2.0,
                        help="Segundos de espera tras aplicar filtros, antes de leer la tabla")
    args = parser.parse_args()
    main(args.url, args.fecha, args.comparador, args.tasa, args.cambio, args.espera)
