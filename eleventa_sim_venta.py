#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
eleventa_sim_venta.py (v3.9.0)

- Ingreso normal (código + ENTER) y Producto Común (Ctrl+P -> Desc -> Cant -> Precio -> ENTER).
- Al finalizar: F12 abrir COBRAR y, si pides --write-notes, F4 para escribir una nota:
  'Codigo Descripcion Cantidad Importe' por renglón + 'Total Cantidad' y 'Total Importe'.

Cambios clave:
- ✅ Regresamos a ingresar SKU repitiendo el CÓDIGO por cada pieza (qty veces).
- ⏱️ Esperas ajustables: after-code-wait (pequeño respiro tras cada captura), between-items (entre piezas),
  after-items-wait (antes de F12). Espera explícita a que aparezca "COBRAR" antes de escribir notas.
- 🧾 Notas consolidadas por (tipo, código, precio) con importe = precio × cantidad; warnings si Excel no cuadra.
"""

import argparse, sys, time, re, unicodedata
from pathlib import Path
from collections import defaultdict

# ---------- Dependencias ----------
try:
    import openpyxl
except Exception:
    print("ERROR: Falta 'openpyxl'. Instala:  pip install openpyxl", file=sys.stderr); raise
try:
    import pyautogui
except Exception:
    pyautogui = None
try:
    import pygetwindow as gw
except Exception:
    gw = None

# pywinauto (win32)
try:
    from pywinauto.application import Application
    from pywinauto.keyboard import send_keys
except Exception:
    Application = None
    send_keys = None

# ---------- Utilidades ----------
def _norm(s: str) -> str:
    if s is None: return ""
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.upper()

def focus_window(window_pattern: str, retries=3, sleep=0.5):
    if gw is None or not window_pattern: return False
    pat = re.compile(re.escape(window_pattern), re.IGNORECASE)
    for _ in range(retries):
        matches = [w for w in gw.getAllWindows() if w.title and pat.search(w.title)]
        if matches:
            win = sorted(matches, key=lambda w: (w.width*w.height), reverse=True)[0]
            try:
                if win.isMinimized: win.restore()
                win.activate(); time.sleep(sleep); return True
            except Exception: time.sleep(sleep)
    return False

def infer_columns(ws):
    header = None
    for r in ws.iter_rows(min_row=1, max_row=5, values_only=True):
        if r and any(isinstance(c, str) and c.strip() for c in r):
            header = [(c.strip().lower() if isinstance(c, str) else "") for c in r]
            break
    idx = {"code": None, "qty": None, "desc": None, "price": None, "importe": None}
    start_row = 1
    if header:
        start_row = 2
        code_names   = {"codigo","código","code","sku","barcode","codbarras","código de barras","cod"}
        qty_names    = {"cantidad","qty","cant","unidades"}
        desc_names   = {"descripcion","descripción","producto","nombre","detalle"}
        price_names  = {"precio usado","precio","valor","monto","price"}
        importe_names= {"importe","total","subtotal"}
        for i,name in enumerate(header):
            if name in code_names    and idx["code"]    is None: idx["code"]    = i
            if name in qty_names     and idx["qty"]     is None: idx["qty"]     = i
            if name in desc_names    and idx["desc"]    is None: idx["desc"]    = i
            if name in price_names   and idx["price"]   is None: idx["price"]   = i
            if name in importe_names and idx["importe"] is None: idx["importe"] = i
    if idx["code"] is None: idx["code"] = 0
    return idx, start_row

def parse_common_code(text: str, common_prefix="COMUN"):
    s = str(text).strip()
    m = re.match(rf"^(?:{re.escape(common_prefix)}|{re.escape(common_prefix)+'-'})(.*)$", s, re.IGNORECASE)
    if not m: return None, None
    tail = m.group(1).strip("-_ ")
    if not tail: return None, None
    parts = re.split(r"[-_]", tail)
    price = None; qty = None
    if parts:
        p = parts[0]
        if re.fullmatch(r"\d+[.,]\d{1,2}", p):
            price = float(p.replace(",", "."))
        elif re.fullmatch(r"\d{3,}", p):
            price = float(p) / 100.0
        elif re.fullmatch(r"\d+", p) and len(parts) == 1:
            qty = int(p)
        if len(parts) >= 2 and qty is None and re.fullmatch(r"\d+", parts[1]):
            qty = int(parts[1])
    return price, qty

# ---------- win32 ----------
def win32_connect(window_pattern: str):
    if Application is None:
        print("pywinauto no disponible; instala: pip install pywinauto", file=sys.stderr); return None, None
    try:
        app = Application(backend="win32").connect(title_re=re.compile(re.escape(window_pattern), re.I))
        win = app.top_window()
        return app, win
    except Exception:
        return None, None

def win32_find_edit(win):
    try:
        for c in win.descendants():
            cls = ""; raw = ""
            try: cls = c.friendly_class_name()
            except Exception: pass
            try: raw = c.class_name()
            except Exception: pass
            if cls == "Edit" or raw in ("TEdit","Edit"):
                return c
    except Exception:
        pass
    return None

def _robust_enter_with(target, text: str, interval=0.0):
    try:
        target.type_keys(text + "{ENTER}", with_spaces=True, pause=interval); return True
    except Exception: pass
    try:
        target.type_keys(text, with_spaces=True, pause=interval)
        if send_keys: send_keys("{ENTER}"); return True
    except Exception: pass
    if pyautogui:
        try:
            pyautogui.write(str(text)); pyautogui.press("enter"); return True
        except Exception: pass
    return False

def win32_type_code(win, code: str, interval=0.0):
    if win is None: return False
    try:
        win.set_focus()
        ed = win32_find_edit(win)
        if ed:
            ed.set_focus()
            try: ed.click_input()
            except Exception: pass
            return _robust_enter_with(ed, code, interval)
        else:
            return _robust_enter_with(win, code, interval)
    except Exception as e:
        print(f"[win32] Error tipeando código: {e}", file=sys.stderr)
        return False

def wait_top_contains(app, keyword_norm, timeout=3.0):
    t0 = time.time()
    kw = _norm(keyword_norm)
    while time.time() - t0 < timeout:
        try:
            top = app.top_window()
            if kw in _norm(top.window_text() or ""): return top
        except Exception: pass
        time.sleep(0.05)
    return None

def win32_add_common(app, win, desc: str, qty: int, price: float, interval=0.0, debug=False):
    """Abre Producto Común (^P), escribe desc, cant, precio y espera volver a la venta."""
    try:
        win.set_focus()
        if send_keys: send_keys("^p")
        else: return False

        dlg = wait_top_contains(app, "PRODUCTO COMUN", timeout=2.0)
        if debug and dlg: print(f"[COMMON] Dialog: {dlg.window_text()!r}")
        target = dlg if dlg is not None else win

        desc_clean = (desc or "").strip().strip(" -–—") or "Producto Comun"
        _robust_enter_with(target, desc_clean, interval)
        _robust_enter_with(target, str(int(qty)), interval)
        price_str = f"{float(price):.2f}".replace(",", ".")
        _robust_enter_with(target, price_str, interval)

        wait_top_contains(app, "VENTA DE PRODUCTOS", timeout=2.0) or \
        wait_top_contains(app, "VENTA", timeout=2.0)
        return True
    except Exception as e:
        print(f"[win32] Error en Producto Común: {e}", file=sys.stderr)
        return False

def win32_write_notes(app, note_lines, interval=0.0, debug=False):
    # esperar COBRAR visible
    cobrar = wait_top_contains(app, "COBRAR", timeout=5.0)
    if debug and cobrar: print(f"[NOTES] Top al cobrar: {cobrar.window_text()!r}")
    if send_keys is None and pyautogui is None: return False

    if send_keys: send_keys("{F4}")
    elif pyautogui: pyautogui.press("f4")
    time.sleep(0.25)

    for line in note_lines:
        try:
            if send_keys:
                send_keys(line, with_spaces=True, pause=interval); send_keys("{ENTER}")
            elif pyautogui:
                pyautogui.write(line); pyautogui.press("enter")
        except Exception:
            pass
        time.sleep(0.03)

    try:
        if send_keys: send_keys("{ENTER}")
        elif pyautogui: pyautogui.press("enter")
    except Exception:
        pass
    return True

# ---------- Notas ----------
def build_notes(items, debug=False):
    for it in items:
        pr = it.get("price")
        im = it.get("importe")
        if isinstance(pr, (int, float)) and isinstance(im, (int, float)):
            esperado = float(pr) * int(it["qty"])
            if abs(esperado - float(im)) > 0.01:
                print(f"[WARN] Importe Excel no coincide para {it['code']}: "
                      f"qty={it['qty']} price={pr} -> {esperado:.2f} vs Excel={im:.2f}",
                      file=sys.stderr)

    agg = defaultdict(lambda: {"qty": 0, "price": None, "desc": ""})
    for it in items:
        key = (it["kind"], it["code"], float(it.get("price") or 0.0))
        agg[key]["qty"]  += int(it["qty"])
        if agg[key]["price"] is None: agg[key]["price"] = it.get("price")
        if not agg[key]["desc"]:      agg[key]["desc"]  = (it.get("desc") or "")

    def _clean_desc(s):
        return re.sub(r"\s+", " ", (s or "").replace("\t"," ")).strip().strip(" -–—")

    note_lines = ["Codigo Descripcion Cantidad Importe"]
    total_qty = 0
    total_importe = 0.0

    for (kind, code, price), v in agg.items():
        desc = _clean_desc(v["desc"]) or "Producto"
        qty  = int(v["qty"])
        if isinstance(price, (int, float)) and price > 0:
            imp = float(price) * qty
        else:
            imp = 0.0
            for it in items:
                if (it["kind"], it["code"], float(it.get("price") or 0.0)) == (kind, code, price):
                    if isinstance(it.get("importe"), (int, float)):
                        imp += float(it["importe"])
        note_lines.append(f"{code} {desc} {qty} {imp:.2f}")
        total_qty     += qty
        total_importe += imp

    note_lines.append(f"Total Cantidad: {total_qty}")
    note_lines.append(f"Total Importe: {total_importe:.2f}")
    if debug:
        print(f"[NOTES] Consolidado renglones: {len(agg)}  Total qty={total_qty}  Total=${total_importe:.2f}")
    return note_lines, total_qty, total_importe

# ---------- Main ----------
def main():
    p = argparse.ArgumentParser(description="Simula la captura de una venta en Eleventa leyendo un Excel (.xlsx).")
    p.add_argument("xlsx_path", nargs="?", help="Ruta al Excel (.xlsx) con códigos y cantidades.")
    p.add_argument("--window", default="eleventa", help="Texto del título de Eleventa (subcadena).")
    p.add_argument("--method", choices=["win32","uia","sendinput"], default="win32")
    p.add_argument("--mode", choices=["repeat","star"], default="repeat")
    p.add_argument("--start-delay", type=float, default=5.0)
    p.add_argument("--between-items", type=float, default=0.25, help="Pausa entre piezas (código por pieza).")
    p.add_argument("--after-code-wait", type=float, default=0.06, help="Micro-pausa después de cada ENTER.")
    p.add_argument("--finalize", action="store_true")
    p.add_argument("--write-notes", action="store_true", help="Tras F12, presiona F4 y escribe notas con el detalle.")
    p.add_argument("--list-windows", action="store_true")
    p.add_argument("--show-pos", action="store_true")
    p.add_argument("--debug", action="store_true")
    p.add_argument("--common-prefix", default="COMUN")
    p.add_argument("--activation-retries", type=int, default=4)
    p.add_argument("--interval", type=float, default=0.0)
    p.add_argument("--after-items-wait", type=float, default=0.9,
                   help="Tiempo extra (s) para que Eleventa termine de actualizar antes de F12.")
    args = p.parse_args()

    if args.list_windows:
        if gw:
            titles = [w.title for w in gw.getAllWindows() if w.title]
            for i,t in enumerate(sorted(set(titles)),1): print(f"{i:02d}  {t}")
        else:
            print("pygetwindow no disponible.")
        sys.exit(0)

    if args.show_pos:
        if pyautogui is None:
            print("pyautogui no disponible; instala: pip install pyautogui", file=sys.stderr)
        else:
            print("Mostrando posición del mouse cada 0.2s (Ctrl+C para salir)...")
            try:
                while True:
                    x, y = pyautogui.position()
                    print(f"X={x:<4}  Y={y:<4}", end="\r", flush=True)
                    time.sleep(0.2)
            except KeyboardInterrupt:
                print("\nOK\n")
        sys.exit(0)

    focus_window(args.window, retries=args.activation_retries)

    if not args.xlsx_path:
        print("ERROR: Indica Excel (.xlsx).", file=sys.stderr); sys.exit(1)
    xlsx = Path(args.xlsx_path)
    if not xlsx.exists():
        print(f"ERROR: No se encontró el archivo: {xlsx}", file=sys.stderr); sys.exit(1)

    wb = openpyxl.load_workbook(xlsx, data_only=True); ws = wb.active
    idx, start_row = infer_columns(ws)

    # Cargar filas
    items = []
    for row in ws.iter_rows(min_row=start_row, values_only=True):
        if not row: continue
        code_cell = row[idx["code"]] if idx["code"] is not None and idx["code"] < len(row) else None
        if code_cell is None or str(code_cell).strip() == "": continue
        code = str(code_cell).strip()

        desc = None; price = None; qty = 1; importe = None; qty_from_excel = False

        if idx["desc"] is not None and idx["desc"] < len(row):
            d = row[idx["desc"]]; desc = (str(d).strip() if d is not None else None)

        if idx["price"] is not None and idx["price"] < len(row):
            pr = row[idx["price"]]
            try:
                if isinstance(pr, str): pr = pr.replace("$","").replace(",","").strip()
                price = float(pr)
            except Exception: price = None

        if idx["importe"] is not None and idx["importe"] < len(row):
            im = row[idx["importe"]]
            try:
                if isinstance(im, str): im = im.replace("$","").replace(",","").strip()
                importe = float(im)
            except Exception: importe = None

        if idx["qty"] is not None and idx["qty"] < len(row):
            q = row[idx["qty"]]
            try:
                if q is not None and str(q).strip() != "":
                    qty = int(float(q));  qty = qty if qty>0 else 1
                    qty_from_excel = True
            except Exception: qty = 1

        is_common = bool(re.match(rf"^{re.escape(args.common_prefix)}(?:$|[-_])", code, re.IGNORECASE))
        if is_common:
            p2, q2 = parse_common_code(code, args.common_prefix)
            if price is None and p2 is not None: price = p2
            if (not qty_from_excel or qty <= 0) and q2 is not None: qty = q2
            if not desc: desc = args.common_prefix
            if importe is None and price is not None: importe = price * qty
            items.append({"kind":"common","code":code,"desc":desc,"qty":qty,"price":price,"importe":importe})
        else:
            if importe is None and price is not None: importe = price * qty
            items.append({"kind":"sku","code":code,"desc":desc,"qty":qty,"price":price,"importe":importe})

    if not items:
        print("No se detectaron artículos en el Excel.", file=sys.stderr); sys.exit(1)

    print(f"Se cargarán {len(items)} artículos. Modo: {args.mode}. Inicio en {args.start_delay:.1f}s...")
    for i in range(int(args.start_delay),0,-1): print(f"... {i}"); time.sleep(1)

    # Ejecutar (win32)
    app, win = win32_connect(args.window)
    if win is None:
        print("No encontré la ventana por Win32. Ajusta --window o ejecuta como admin si Eleventa lo está.", file=sys.stderr); sys.exit(2)

    # Ingreso de artículos (SKU: repetir código por pieza)
    for idx_i, it in enumerate(items, 1):
        if args.debug:
            print(f"[{idx_i}] {it['kind'].upper()} -> {it['code']}  qty={it['qty']}  desc={it.get('desc')}  price={it.get('price')}  imp={it.get('importe')}")
        if it["kind"] == "common":
            ok = win32_add_common(app, win, it["desc"], it["qty"], (it["price"] or 0.0), interval=args.interval, debug=args.debug)
            if not ok: print(f"[WARN] No se pudo ingresar Producto Común para {it['code']}.", file=sys.stderr)
        else:
            reps = max(1, int(it["qty"]))
            for _ in range(reps):
                ok = win32_type_code(win, it["code"], args.interval)
                if not ok:
                    print(f"[WARN] No se pudo tipear código {it['code']}. Reintentando...", file=sys.stderr)
                    time.sleep(0.15)
                    win32_type_code(win, it["code"], args.interval)
                time.sleep(args.after_code_wait)
            time.sleep(args.between_items)

    # Notas consolidadas
    note_lines, total_qty, total_importe = build_notes(items, debug=args.debug)

    # Finalizar + Notas
    if args.finalize:
        time.sleep(max(0.0, args.after_items_wait))
        try:
            if send_keys: send_keys("{F12}")
            elif pyautogui: pyautogui.press("f12")
        except Exception:
            pass
        # Esperar a que COBRAR esté visible antes de F4/notas
        wait_top_contains(app, "COBRAR", timeout=5.0)
        time.sleep(0.25)
        if args.write_notes:
            win32_write_notes(app, note_lines, interval=args.interval, debug=args.debug)

    print("Listo. Notas generadas" + (" y escritas." if args.finalize and args.write_notes else "."))

if __name__ == "__main__":
    main()
