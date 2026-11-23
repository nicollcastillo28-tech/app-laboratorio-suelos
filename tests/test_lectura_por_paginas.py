"""Supabase entrega como máximo 1000 filas por consulta: las listas completas se leen por páginas (si no, con más de 1000 ensayos los más
nuevos no se veían y al abrirlos la app intentaba crearlos otra vez: error de la base por el duplicado)."""
import os
import sys
import types
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import db  # noqa: E402
import test_guardado_ensayo as falsa  # noqa: E402


class QConTope(falsa.Q):
    """Como la base real: sin .range(), o con un rango más grande, nunca más de 1000 filas."""
    def execute(s):
        res = super().execute()
        if s.ins is None and s.upd is None:
            res = types.SimpleNamespace(data=res.data[:1000])
        return res


class ClienteConTope:
    def table(self, nombre):
        return QConTope(nombre)


class DuplicadoQ(QConTope):
    def execute(s):
        if s.ins is not None and any(r["muestra_id"] == s.ins["muestra_id"] and r["tipo"] == s.ins["tipo"] and not r.get("archived")
                                     for r in falsa.TABLAS["assays"]):
            raise RuntimeError("duplicate key value violates unique constraint \"assays_muestra_tipo_uk\"")
        return super().execute()


class ClienteDuplicado:
    def table(self, nombre):
        return DuplicadoQ(nombre)


class LecturaPorPaginasTest(unittest.TestCase):
    def setUp(self):
        for t in falsa.TABLAS:
            falsa.TABLAS[t] = []
        db._COLUMNAS_LIVIANAS = None
        db._CACHE_PESADO.clear()
        self._cliente = db.get_client
        db.get_client = lambda: ClienteConTope()

    def tearDown(self):
        db.get_client = self._cliente

    def test_mas_de_1000_ensayos_se_leen_todos(self):
        falsa.TABLAS["assays"] = [{"id": f"a{i:05d}", "muestra_id": f"m{i}", "tipo": "humedad", "archived": False, "data": {},
                                   "historial": [], "created_at": f"2026-01-01T00:00:{i:06d}", "updated_at": "x"} for i in range(2345)]
        self.assertEqual(len(db.list_all_assays()), 2345)
        self.assertEqual(len(db.list_all_assays()), 2345)   # segunda vez: solo columnas livianas
        falsa.TABLAS["muestras"] = [{"id": f"m{i}", "numero": str(i), "archived": False} for i in range(1500)]
        self.assertEqual(len(db.list_all_muestras()), 1500)

    def test_abrir_un_ensayo_que_ya_existe_devuelve_el_existente(self):
        falsa.TABLAS["assays"] = [{"id": "a1", "muestra_id": "m1", "tipo": "humedad", "archived": False, "data": {"x": 1}}]
        db.get_client = lambda: ClienteDuplicado()
        self.assertEqual(db.create_assay("m1", "humedad")["id"], "a1")
        self.assertEqual(len(falsa.TABLAS["assays"]), 1)
        self.assertEqual(db.create_assay("m2", "humedad")["muestra_id"], "m2")


if __name__ == "__main__":
    unittest.main()
