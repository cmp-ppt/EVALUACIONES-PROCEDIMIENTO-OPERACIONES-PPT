# -*- coding: utf-8 -*-
"""Extrae trabajadores y procedimientos desde index.html y el cruce de RUT."""
import re, json, os

_AQUI = os.path.dirname(os.path.abspath(__file__))

src = open(os.path.join(os.path.dirname(_AQUI), 'index.html'), encoding='utf-8').read()

def bloque(marca, largo):
    i = src.index(marca)
    return src[i:i+largo].replace('\\n', '\n')

# --- procedimientos ---
procs = []
for code, nombre in re.findall(r"\{ code:'([\d.]+)',\s*sub:'[\d.]+',\s*nombre:'([^']+)'", bloque('PROCS() {', 4000)):
    procs.append({'code': code, 'nombre': nombre})

# --- trabajadores (mismo orden y mismos ids que genera la app: k arranca en 0) ---
people, k = [], 0
for g, rows in re.findall(r"\{ g:'(G\d)', rows:\[(.*?)\]\s*\}", bloque('RAW_PEOPLE() {', 4300), re.S):
    for full, cargo in re.findall(r"\['([^']+)','([^']+)'\]", rows):
        ap_part, nom_part = [s.strip() for s in full.split(',')]
        people.append({'id': 'p%d' % k, 'g': g, 'apellidos': ap_part,
                       'nombres': nom_part, 'cargo': cargo})
        k += 1

# los borrados del dashboard
people = [p for p in people if p['id'] not in ('p8', 'p36')]

# --- RUT: guardados junto a este modulo, para no depender de archivos externos ---
_aqui = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(_aqui, 'ruts.json'), encoding='utf-8') as _fh:
    ruts = json.load(_fh)
for p in people:
    p['rut'] = ruts.get(p['id'], 'POR CONFIRMAR')

orden_cargo = {'Supervisor de Operaciones': 0, 'Controlador Operación Puerto': 1,
               'Operador Líder Puerto': 2, 'Operador Especialista Puerto': 3,
               'Operador Avanzado Puerto': 4, 'Operador Base': 5}
people.sort(key=lambda p: (p['g'], orden_cargo.get(p['cargo'], 9), p['apellidos']))

if __name__ == '__main__':
    print(len(procs), 'procedimientos |', len(people), 'trabajadores')
    print('sin rut:', [p['id'] + ' ' + p['apellidos'] for p in people if p['rut'] == 'POR CONFIRMAR'])
    print(procs[0], procs[-1])
